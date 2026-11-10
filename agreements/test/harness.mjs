// Local test harness: runs the real Code.gs with small fakes of the Google services and
// serves docs/ and the web app API on http://127.0.0.1:18790.
//
//   node agreements/test/harness.mjs
//
// Emails are kept in memory (GET /admin/outbox). To send them through a real relay instead,
// set RELAY_URL and RELAY_KEY, for example to `npx wrangler dev --port 18787` in mailer/.
//
// What Sophie sees and does:
//   http://127.0.0.1:18790/admin/form              the New agreement form
//   http://127.0.0.1:18790/admin/form?copyRow=2    the form, copied from row 2
//   http://127.0.0.1:18790/b/?CODE                 the client page
// Other admin actions:
//   curl -X POST 'http://127.0.0.1:18790/admin/add' -d '{"Client name":"Jane Smith",...}'
//   curl -X POST 'http://127.0.0.1:18790/admin/send?row=2'
//   curl 'http://127.0.0.1:18790/admin/rows'
import { createHash, randomUUID } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync, existsSync, statSync } from "node:fs";
import { createServer } from "node:http";
import { dirname, extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const here = dirname(fileURLToPath(import.meta.url));
const docs = join(here, "..", "..", "docs");
const PORT = 18790;

// --- Fake Google services -------------------------------------------------

class FakeRange {
  constructor(sheet, row, col, rows, cols) {
    Object.assign(this, { sheet, row, col, rows, cols });
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.rows; r++) {
      const line = [];
      for (let c = 0; c < this.cols; c++)
        line.push(this.sheet.cell(this.row + r, this.col + c));
      out.push(line);
    }
    return out;
  }
  setValue(v) {
    this.sheet.write(this.row, this.col, v);
    return this;
  }
  setValues(values) {
    values.forEach((line, r) =>
      line.forEach((v, c) => this.sheet.write(this.row + r, this.col + c, v)),
    );
    return this;
  }
  // The fake keeps every value as given, which is what plain text format gives in Sheets.
  setNumberFormat() {
    return this;
  }
  setFontWeight() {
    return this;
  }
  getRow() {
    return this.row;
  }
  protect() {
    this.sheet.protections.push(this.row);
    const p = { setDescription: () => p, setWarningOnly: () => p };
    return p;
  }
}

class FakeSheet {
  constructor(name) {
    this.name = name;
    this.data = [];
    this.protections = [];
    this.activeRow = 1;
  }
  getName() {
    return this.name;
  }
  cell(r, c) {
    const v = (this.data[r - 1] || [])[c - 1];
    return v === undefined ? "" : v;
  }
  write(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    this.data[r - 1][c - 1] = v;
  }
  getLastRow() {
    return this.data.length;
  }
  getLastColumn() {
    return Math.max(0, ...this.data.map((r) => r.length));
  }
  appendRow(values) {
    this.data.push([...values]);
  }
  getRange(row, col, rows = 1, cols = 1) {
    // A1 notation such as "A:Q" is only used to set formats, which the fake ignores.
    if (typeof row === "string") return new FakeRange(this, 1, 1, 0, 0);
    return new FakeRange(this, row, col, rows, cols);
  }
  getDataRange() {
    return new FakeRange(this, 1, 1, this.getLastRow(), this.getLastColumn());
  }
  getActiveRange() {
    return new FakeRange(this, this.activeRow, 1, 1, 1);
  }
  setFrozenRows() {}
}

const sheets = {};
const spreadsheet = {
  getSheetByName: (n) => sheets[n] || null,
  insertSheet: (n) => (sheets[n] = new FakeSheet(n)),
};
const props = {};
const drive = [];
const alerts = [];
const dialogs = [];
const outbox = [];
const useRealRelay = Boolean(process.env.RELAY_URL);

const google = {
  SpreadsheetApp: {
    getActiveSpreadsheet: () => spreadsheet,
    getActiveSheet: () => sheets.Agreements,
    getUi: () => ({
      alert: (m) => alerts.push(m),
      showModalDialog: (output) => dialogs.push(output),
      createMenu: () => ({
        addItem() {
          return this;
        },
        addSeparator() {
          return this;
        },
        addToUi() {},
      }),
    }),
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null),
      setProperty: (k, v) => (props[k] = v),
    }),
  },
  UrlFetchApp: {
    fetch(url, opts) {
      if (!useRealRelay) {
        outbox.push(JSON.parse(opts.payload));
        return {
          getResponseCode: () => 200,
          getContentText: () => '{"ok":true}',
        };
      }
      const args = [
        "-s",
        "-o",
        "-",
        "-w",
        "\n%{http_code}",
        "-X",
        "POST",
        url,
        "-H",
        `Content-Type: ${opts.contentType}`,
      ];
      for (const [k, v] of Object.entries(opts.headers || {}))
        args.push("-H", `${k}: ${v}`);
      args.push("--data-binary", "@-");
      const out = execFileSync("curl", args, {
        input: opts.payload,
      }).toString();
      const i = out.lastIndexOf("\n");
      return {
        getResponseCode: () => Number(out.slice(i + 1)),
        getContentText: () => out.slice(0, i),
      };
    },
  },
  Utilities: {
    DigestAlgorithm: { SHA_256: "sha256" },
    computeDigest: (alg, s) =>
      [...createHash(alg).update(s).digest()].map((b) =>
        b > 127 ? b - 256 : b,
      ),
    getUuid: () => randomUUID(),
    formatDate: (d) => {
      const p = Object.fromEntries(
        new Intl.DateTimeFormat("en-GB", {
          timeZone: "Europe/London",
          year: "numeric",
          month: "2-digit",
          day: "2-digit",
          hour: "2-digit",
          minute: "2-digit",
          second: "2-digit",
          hour12: false,
        })
          .formatToParts(d)
          .map((x) => [x.type, x.value]),
      );
      return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`;
    },
    // The real service renders a PDF. Here the "PDF" holds the HTML so tests can read it.
    newBlob: (content) => {
      const blob = {
        name: "",
        getAs: () => blob,
        setName: (n) => ((blob.name = n), blob),
        getBytes: () => [...Buffer.from(content)],
        html: content,
      };
      return blob;
    },
    base64Encode: (bytes) =>
      Buffer.from(bytes.map((b) => b & 255)).toString("base64"),
  },
  DriveApp: {
    createFolder: () => ({
      getId: () => "folder1",
      createFile: (blob) => {
        drive.push(blob);
        return { getUrl: () => `https://drive.example/file/${drive.length}` };
      },
    }),
    getFolderById: () => google.DriveApp.createFolder(),
  },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  ContentService: {
    MimeType: { JSON: "json" },
    createTextOutput: (s) => ({
      content: s,
      setMimeType() {
        return this;
      },
    }),
  },
  HtmlService: {
    // Form.html is rendered by GET /admin/form, which reads prefillJson from the dialog.
    createTemplateFromFile: () => ({
      evaluate() {
        const out = {
          prefillJson: this.prefillJson,
          setWidth: () => out,
          setHeight: () => out,
        };
        return out;
      },
    }),
    createHtmlOutput: () => ({
      setWidth() {
        return this;
      },
      setHeight() {
        return this;
      },
    }),
  },
  Session: { getActiveUser: () => ({ getEmail: () => "owner@example.com" }) },
  console,
};

const context = vm.createContext({ ...google });
vm.runInContext(
  readFileSync(join(here, "..", "apps-script", "Code.gs"), "utf8"),
  context,
  { filename: "Code.gs" },
);

// --- Set up the sheet the same way the "Set up" menu does (minus the prompts) ---

Object.assign(props, {
  RELAY_URL: process.env.RELAY_URL || "http://127.0.0.1:18787",
  RELAY_KEY: process.env.RELAY_KEY || "test-key",
  NOTIFY_EMAIL: "imperialhomes@btinternet.com",
  SITE_URL: `http://127.0.0.1:${PORT}`,
});
const ui = google.SpreadsheetApp.getUi();
google.SpreadsheetApp.getUi = () => ({
  ...ui,
  prompt: () => ({ getSelectedButton: () => "OK", getResponseText: () => "" }),
  Button: { OK: "OK" },
  ButtonSet: {},
});
context.setup();
sheets.Terms.data[1] = [
  "v1",
  true,
  "1. Test terms.\n2. The success fee is payable on completion.",
];
const bank = sheets.Template.data.find((r) => r[0] === "Bank details");
bank[1] =
  "Account name: Imperial Homes Property Search\nSort code: 11-22-33\nAccount number: 12345678";

// --- HTTP server ------------------------------------------------------------

const types = {
  ".html": "text/html",
  ".css": "text/css",
  ".js": "text/javascript",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

function send(res, status, body, type = "application/json") {
  res.writeHead(status, {
    "Content-Type": type,
    "Access-Control-Allow-Origin": "*",
  });
  res.end(
    typeof body === "string" || Buffer.isBuffer(body)
      ? body
      : JSON.stringify(body),
  );
}

// Stands in for google.script.run inside Form.html: calls go to POST /admin/call.
const GOOGLE_SCRIPT_SHIM = `<script>
function runner(ok, fail) {
  return new Proxy({}, { get: function (_, fn) {
    if (fn === 'withSuccessHandler') return function (f) { return runner(f, fail); };
    if (fn === 'withFailureHandler') return function (f) { return runner(ok, f); };
    return function () {
      fetch('/admin/call', { method: 'POST', body: JSON.stringify({ fn: fn, args: [].slice.call(arguments) }) })
        .then(function (r) { return r.json(); })
        .then(function (d) { if (d.error) { fail && fail(new Error(d.error)); } else { ok && ok(d.result); } });
    };
  } });
}
window.google = { script: { run: runner(), host: { close: function () { document.body.textContent = 'Closed'; } } } };
</script>`;

function rowsView() {
  const s = sheets.Agreements;
  const headers = s.data[0];
  return s.data.slice(1).map((r, i) => ({
    row: i + 2,
    ...Object.fromEntries(headers.map((h, j) => [h, r[j] ?? ""])),
  }));
}

createServer(async (req, res) => {
  const url = new URL(req.url, `http://127.0.0.1:${PORT}`);
  let body = "";
  for await (const chunk of req) body += chunk;
  try {
    if (url.pathname === "/api" && req.method === "GET") {
      return send(
        res,
        200,
        context.doGet({ parameter: Object.fromEntries(url.searchParams) })
          .content,
      );
    }
    if (url.pathname === "/api" && req.method === "POST") {
      return send(
        res,
        200,
        context.doPost({ postData: { contents: body } }).content,
      );
    }
    if (url.pathname === "/admin/add") {
      const values = JSON.parse(body);
      sheets.Agreements.appendRow(
        sheets.Agreements.data[0].map((h) => values[h] ?? ""),
      );
      return send(res, 200, { row: sheets.Agreements.data.length });
    }
    if (url.pathname === "/admin/send" || url.pathname === "/admin/cancel") {
      sheets.Agreements.activeRow = Number(url.searchParams.get("row"));
      alerts.length = 0;
      context[
        url.pathname === "/admin/send" ? "sendSelected" : "cancelSelected"
      ]();
      return send(res, 200, { alerts });
    }
    if (url.pathname === "/admin/rows") return send(res, 200, rowsView());
    if (url.pathname === "/admin/pdf")
      return send(
        res,
        200,
        drive[Number(url.searchParams.get("n")) - 1].html,
        "text/html",
      );
    if (url.pathname === "/admin/protected")
      return send(res, 200, sheets.Agreements.protections);
    if (url.pathname === "/admin/outbox") {
      if (req.method === "DELETE") outbox.length = 0;
      return send(res, 200, outbox);
    }
    if (url.pathname === "/admin/form") {
      dialogs.length = 0;
      alerts.length = 0;
      const copyRow = url.searchParams.get("copyRow");
      if (copyRow) sheets.Agreements.activeRow = Number(copyRow);
      context[copyRow ? "openCopyOfSelected" : "openNewAgreement"]();
      if (!dialogs.length) return send(res, 400, { alerts });
      const html = readFileSync(
        join(here, "..", "apps-script", "Form.html"),
        "utf8",
      )
        .replace("<?!= prefillJson ?>", () => dialogs[0].prefillJson)
        .replace("<body>", () => `<body>${GOOGLE_SCRIPT_SHIM}`);
      return send(res, 200, html, "text/html");
    }
    if (url.pathname === "/admin/call") {
      const { fn, args } = JSON.parse(body);
      if (!["createAgreement", "sendDraft"].includes(fn))
        return send(res, 400, { error: `not allowed: ${fn}` });
      try {
        return send(res, 200, { result: context[fn](...args) });
      } catch (err) {
        return send(res, 200, { error: err.message });
      }
    }

    let path = normalize(join(docs, decodeURIComponent(url.pathname)));
    if (!path.startsWith(docs))
      return send(res, 403, "forbidden", "text/plain");
    if (existsSync(path) && statSync(path).isDirectory())
      path = join(path, "index.html");
    if (!existsSync(path)) return send(res, 404, "not found", "text/plain");
    return send(
      res,
      200,
      readFileSync(path),
      types[extname(path)] || "application/octet-stream",
    );
  } catch (err) {
    console.error(err);
    return send(res, 500, { error: String(err.message || err) });
  }
}).listen(PORT, "127.0.0.1", () =>
  console.log(`harness ready on http://127.0.0.1:${PORT}`),
);
