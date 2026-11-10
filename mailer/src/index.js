// Mail relay: the Google Apps Script backend cannot speak SMTP, and btinternet.com
// rejects mail (DMARC p=reject) unless it leaves through BT's own server.
// This Worker takes one JSON message and sends it through BT SMTP. Nothing is stored.
import { WorkerMailer } from "worker-mailer";

const MAX_RECIPIENTS = 5;
const MAX_ATTACHMENT_BYTES = 5 * 1024 * 1024;
const EMAIL = /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/;

export default {
  async fetch(request, env) {
    if (request.method !== "POST") return json({ error: "POST only" }, 405);

    const auth = request.headers.get("Authorization") || "";
    if (!env.RELAY_KEY || !safeEqual(auth, `Bearer ${env.RELAY_KEY}`)) {
      return json({ error: "unauthorized" }, 401);
    }

    let msg;
    try {
      msg = await request.json();
    } catch {
      return json({ error: "body must be JSON" }, 400);
    }
    const problem = validate(msg);
    if (problem) return json({ error: problem }, 400);

    const port = Number(env.SMTP_PORT || 465);
    try {
      await WorkerMailer.send(
        {
          host: env.SMTP_HOST,
          port,
          secure: port === 465,
          startTls: port !== 465,
          credentials: { username: env.SMTP_USER, password: env.SMTP_PASS },
          authType: ["plain", "login"],
          socketTimeoutMs: 20000,
          responseTimeoutMs: 20000,
        },
        {
          from: { name: env.FROM_NAME || "", email: env.SMTP_USER },
          to: msg.to,
          bcc: msg.bcc,
          reply: msg.replyTo,
          subject: msg.subject,
          text: msg.text,
          html: msg.html,
          attachments: msg.attachments,
        },
      );
    } catch (err) {
      return json(
        { error: `smtp: ${err && err.message ? err.message : String(err)}` },
        502,
      );
    }
    return json({ ok: true });
  },
};

function validate(msg) {
  if (!msg || typeof msg !== "object") return "body must be an object";
  const to = [].concat(msg.to || []);
  const bcc = [].concat(msg.bcc || []);
  if (to.length === 0) return "to is required";
  if (to.length + bcc.length > MAX_RECIPIENTS)
    return `at most ${MAX_RECIPIENTS} recipients`;
  for (const addr of [...to, ...bcc]) {
    if (typeof addr !== "string" || !EMAIL.test(addr))
      return `bad address: ${addr}`;
  }
  if (msg.replyTo && !EMAIL.test(msg.replyTo)) return "bad replyTo";
  if (
    typeof msg.subject !== "string" ||
    !msg.subject ||
    msg.subject.length > 300
  )
    return "bad subject";
  if (!msg.text && !msg.html) return "text or html is required";
  const attachments = msg.attachments || [];
  if (!Array.isArray(attachments) || attachments.length > 3)
    return "at most 3 attachments";
  let bytes = 0;
  for (const a of attachments) {
    if (!a || typeof a.filename !== "string" || typeof a.content !== "string")
      return "bad attachment";
    bytes += a.content.length * 0.75;
  }
  if (bytes > MAX_ATTACHMENT_BYTES) return "attachments too large";
  return null;
}

// Constant-time compare so the key cannot be guessed from response timing.
function safeEqual(a, b) {
  const x = new TextEncoder().encode(a);
  const y = new TextEncoder().encode(b);
  let diff = x.length ^ y.length;
  for (let i = 0; i < Math.max(x.length, y.length); i++)
    diff |= (x[i] || 0) ^ (y[i] || 0);
  return diff === 0;
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}
