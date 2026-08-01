# Office 365 Playwright Check

This project automates a basic Microsoft 365 login check with Playwright.

## Features

- Logs in with a single Office 365 account
- Supports batch login from a credentials file
- Uses optional SOCKS5 proxy settings from a local or repo-level proxy file
- Prints compact status results such as `[OK]`, `[FAIL]`, and `[2FA]`
- Includes a GitHub Actions workflow for automated runs

## Requirements

- Node.js 20+
- npm

## Install

```bash
npm install
npx playwright install --with-deps chromium
```

## Usage

### Single account

Set environment variables:

```bash
export OFFICE365_EMAIL="your-email@example.com"
export OFFICE365_PASSWORD="your-password"
```

Run:

```bash
npm run check
```

### Batch accounts from file

Create a file named `accounts.txt` with lines in this format:

```txt
email@example.com|password123
```

Run:

```bash
npm run login-file
```

### Proxy file

Add one proxy per line to `proxy.txt` or to the local Mac file:

```txt
host:port:user:pass
```

The script will read the local file at `~/.office365-playwright/proxy.txt` first if present.

## GitHub Actions

The repository includes a workflow in `.github/workflows/office365-check.yml`.

Add these repository secrets before running it:

- `OFFICE365_EMAIL`
- `OFFICE365_PASSWORD`
- `OFFICE365_MFA_CODE`

## Notes

- The repository ignores sensitive files such as `accounts.txt`, `proxy.txt`, and `.env`.
- For private or organization use, keep credentials in local files or GitHub Secrets.
