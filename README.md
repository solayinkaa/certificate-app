# Certificate Distribution

A Streamlit app that stamps each recipient's name and certificate number onto a
template (PDF, PNG or JPG) and e-mails every personalised PDF from an Excel
list. All processing happens in memory, so concurrent users never share files.

## Repository layout

```
app.py                         # The whole application
requirements.txt               # Pinned dependencies
.streamlit/config.toml         # Server settings (iframe-friendly uploads)
.streamlit/secrets.toml.example  # Template for local secrets (not deployed)
.gitignore
README.md
```

No `runtime.txt` or `packages.txt` is needed. Community Cloud ignores
`runtime.txt` (the Python version is picked at deploy time), and PyMuPDF ships
its own binaries, so no system packages are required.

## 1. Put the code on GitHub

1. Create a **private** repository on GitHub (recommended, see "License keys"
   below) and upload every file above, keeping the `.streamlit` folder.
2. Do **not** commit a real `.streamlit/secrets.toml`. It is already in
   `.gitignore`.

## 2. Deploy on Streamlit Community Cloud

1. Sign in at <https://share.streamlit.io> with GitHub and allow access to
   private repositories if asked.
2. Click **Create app**, then **Deploy a public app from GitHub**.
3. Choose the repository, branch `main` and main file `app.py`.
4. Optionally pick a custom subdomain, e.g. `c2s-certificates`.
5. Open **Advanced settings**:
   - **Python version: choose 3.13** (3.12 is fine too). The app was tested
     on 3.13, and every pinned package also ships ready-made builds for
     3.14. To change the version later you must delete and redeploy the app.
   - **Secrets:** paste your Pro keys (see below).
6. Click **Deploy** and wait for the build to finish.

## 3. License keys (keep them out of the public code)

`VALID_KEYS` in `app.py` is visible to anyone who can read the repository.
Either keep the repository private, or replace those two keys with keys nobody
knows and add your real keys as secrets. The app merges secrets with
`VALID_KEYS` automatically.

In **Advanced settings → Secrets** (or later in **App settings → Secrets**):

```toml
PRO_LICENSE_KEYS = ["PRO-CUSTOMER-001", "PRO-CUSTOMER-002"]
```

A comma-separated string also works: `PRO_LICENSE_KEYS = "KEY-1, KEY-2"`.
Updating secrets takes effect within a minute, without redeploying.

## 4. Embed it on c2snet.org

Add `?embed=true` to the app URL. This hides Streamlit's toolbar and footer:

```html
<iframe
  src="https://c2s-certificates.streamlit.app/?embed=true"
  style="width:100%; height:1400px; border:0;"
  allow="clipboard-write"
  title="Certificate Distribution">
</iframe>
```

Optional extras: `&embed_options=light_theme` forces the light theme, and
`&embed_options=show_padding` keeps the default page padding.

**If uploads fail inside the iframe** with "AxiosError: Request failed with
status code 403", the browser blocked Streamlit's security cookie because the
app sits on a different site. `.streamlit/config.toml` already turns that check
off (`enableXsrfProtection = false`). If the error still appears, open the
app's direct link in a new tab to confirm it works there, then use a
"Open the certificate tool" button that links to the app as a fallback.

Free Community Cloud apps go to sleep after several days without visitors.
The first visitor then sees a "wake up" button, which takes under a minute.

## 5. E-mail settings

### Gmail (recommended for small batches)

1. Turn on **2-Step Verification** for the Google account
   (Google Account → Security).
2. Open <https://myaccount.google.com/apppasswords>, create an app password
   named e.g. "Certificates", and copy the 16-character code.
3. In the app sidebar enter:
   - SMTP Server: `smtp.gmail.com`
   - SMTP Port: `465` (SSL) or `587` (STARTTLS); both work
   - Sender Email: the full Gmail address
   - App Password: the 16-character code (not your normal password)
4. Click **Test SMTP Connection** before sending a batch.

Gmail caps personal accounts at about 500 recipients per day (2,000 for Google
Workspace). Keep the delay slider at 1 second or more for large batches.
Google sometimes blocks sign-ins from cloud servers. When that happens the app
shows an "Authentication failed" message, and the cPanel option below is the
fix.

### cPanel / your own domain (fallback)

1. In cPanel → **Email Accounts**, create a mailbox such as
   `certificates@c2snet.org`, then click **Connect Devices** to see the
   outgoing server details.
2. In the app sidebar enter:
   - SMTP Server: the outgoing host shown there, usually `mail.c2snet.org`
   - SMTP Port: `465`
   - Sender Email: the full mailbox address
   - App Password: the mailbox password
3. If the test fails, ask your host whether external SMTP logins on port 465
   (or 587) are allowed and what the hourly sending limit is, then set the
   delay slider to stay under it.

## 6. Using the app

1. **Step 1:** upload the template, then move the X, Y and size sliders until
   "Jane Doe" and "CERT-001" sit correctly in the preview.
2. **Step 2:** upload an `.xlsx` with columns `Name`, `CertNo` and `Email`.
   The free tier processes the first 50 rows; a Pro key removes the limit.
3. **Step 3:** click **Start Bulk Dispatch**, keep the tab open until it
   finishes, then download the CSV report.

## Run locally

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # optional
streamlit run app.py
```
