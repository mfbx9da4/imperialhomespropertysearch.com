# Client agreements: setup

Sophie sends each client a link to their property search brief, fees and terms.
The client types their name, ticks a box and clicks Accept. The client gets a
confirmation email with a signed PDF, and Imperial Homes gets a copy.

## How it fits together

| Part           | Where                                                                                                           | Cost |
| -------------- | --------------------------------------------------------------------------------------------------------------- | ---- |
| Admin and data | A Google Sheet with the script in `apps-script/`                                                                | Free |
| Client page    | `docs/b/index.html` on this GitHub Pages site, so links look like `imperialhomespropertysearch.co.uk/b/?7KQ2MX` | Free |
| Signed copies  | PDFs in the Google Drive folder "Imperial Homes signed agreements"                                              | Free |
| Email          | `mailer/`, a Cloudflare Worker that sends through BT (`mail.btinternet.com`, port 587) as `imperialhomes@btinternet.com`  | Free |

Why the mail relay: Apps Script cannot send through another provider's mail server, and
btinternet.com tells other mail servers to reject (DMARC `p=reject`) any email that claims to be from
`@btinternet.com` but does not come from BT. So the email must go out through BT's own server.

No passwords or client data are stored in this repository.

## 1. Mail relay (about 5 minutes)

1. Make a free Cloudflare account at https://dash.cloudflare.com/sign-up.
2. In a terminal:

   ```bash
   cd mailer
   npm install
   npx wrangler login
   npx wrangler deploy
   ```

   Copy the `https://imperial-homes-mailer.<name>.workers.dev` URL that `deploy` prints.

3. Save the BT password. The command asks for it, so it is never written to a file:

   ```bash
   npx wrangler secret put SMTP_PASS
   ```

4. Make a random key that lets only the Sheet use the relay. The command shows the key once. Keep it for step 2.4:

   ```bash
   openssl rand -hex 24 | tee /dev/tty | npx wrangler secret put RELAY_KEY
   ```

## 2. Google Sheet (about 10 minutes)

Use a Google account that belongs to Imperial Homes, so the system stays with the business.

1. Make a new Google Sheet called "Imperial Homes Agreements".
2. Open Extensions > Apps Script. Replace everything in `Code.gs` with `apps-script/Code.gs`.
3. In Project Settings, tick "Show appsscript.json manifest file in editor". Replace `appsscript.json` with `apps-script/appsscript.json`. Save.
4. Reload the Sheet. Use the new **Imperial Homes > Set up** menu. Give it the relay URL and key from step 1.
   For the other questions, leave the box blank to keep the default.
5. On the **Terms** tab, paste the terms and conditions into `Text`. On the **Template** tab, replace the
   "Bank details" text with the real account details. The confirmation email includes them so the client can pay
   the deposit. Change any other wording you like. `{Client Name}`, `{First Name}`, `{Link}`, `{Deposit}`,
   `{Success Fee}`, `{Reference}` and `{Bank Details}` are filled in for each client. The script will not send
   an agreement until the terms and bank details are filled in.
6. In Apps Script, click Deploy > New deployment > Web app. Set "Execute as" to Me and "Who has access" to Anyone.
   Click Deploy, allow the permissions, and copy the web app URL.
7. Use **Imperial Homes > Send a test email to me** to check that email works.

## 3. Client page

In `docs/b/index.html`, set `API_URL` to the web app URL from step 2.6. Commit and push. GitHub Pages publishes it.

If you change the script later, use Deploy > Manage deployments > Edit > New version. This keeps the same URL.

## Daily use

1. **Imperial Homes > New agreement** opens a form. The brief starts from the "Brief template" on the
   Template tab, and the fee and deposit start from "Default success fee" and "Default deposit".
2. Most clients are like an earlier one: click that client's row on the **Agreements** tab, then use
   **Imperial Homes > New agreement, copied from selected row**. Type the new name and email and change
   what is different.
3. Click **Preview** to see the exact page the client will see. Accept does not work on a preview.
   Then click **Looks good, send to client**. The client gets an email with their link.
4. The form then shows the link and a **Send on WhatsApp** button, if you want to send it that way too.
5. The Status column changes from `draft` to `sent` to `viewed` to `accepted`. When the client accepts,
   "Signed PDF" links to the signed copy, and the client gets a confirmation email with the bank details
   for the deposit. You get a copy at the NOTIFY_EMAIL address.
6. To withdraw a link, use **Imperial Homes > Cancel agreement**. To change a brief after you sent it,
   cancel it and make a new agreement.

## Local testing

```bash
node agreements/test/harness.mjs
```

`harness.mjs` runs the real `Code.gs` and `Form.html` with fake Google services and keeps emails in memory.
Its header lists the URLs, such as http://127.0.0.1:18790/admin/form for the New agreement form.
To send real emails, run the relay with `npx wrangler dev --port 18787` in `mailer/` and start the
harness with `RELAY_URL=http://127.0.0.1:18787 RELAY_KEY=<key from mailer/.dev.vars>`.
`mailer/save-bt-password.sh` saves the BT password into `mailer/.dev.vars` with typing hidden.
