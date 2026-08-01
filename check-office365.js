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
  const fs = require('fs');
  const path = require('path');
  const localAccountPath = path.join(process.env.HOME || '', '.office365-playwright', 'accounts.txt');
  const repoAccountPath = path.join(__dirname, 'accounts.txt');

  let email = process.env.OFFICE365_EMAIL;
  let password = process.env.OFFICE365_PASSWORD;
  const mfaCode = process.env.OFFICE365_MFA_CODE;
  const headless = process.env.PLAYWRIGHT_HEADLESS !== 'false';

  if (!email || !password) {
    const accountFilePath = fs.existsSync(localAccountPath) ? localAccountPath : repoAccountPath;
    if (fs.existsSync(accountFilePath)) {
      const firstLine = fs.readFileSync(accountFilePath, 'utf8').split(/\r?\n/).find(Boolean);
      if (firstLine) {
        const [fileEmail, filePassword] = firstLine.split('|').map((value) => value.trim());
        if (fileEmail && filePassword) {
          email = fileEmail;
          password = filePassword;
        }
      }
    }
  }

  if (!email || !password) {
    console.error('Harap set environment variable: OFFICE365_EMAIL dan OFFICE365_PASSWORD atau isi accounts.txt');
    process.exit(1);
  }

  const proxyConfigs = getProxyConfigs();
  const proxyConfig = proxyConfigs.length ? buildProxyConfig(proxyConfigs[0]) : null;

  let browser;
  try {
    browser = await chromium.launch({ headless });
    const context = await browser.newContext({
      viewport: { width: 1440, height: 1200 },
      ...(proxyConfig ? { proxy: proxyConfig } : {})
    });
    const page = await context.newPage();

    await page.goto('https://login.microsoftonline.com/', { waitUntil: 'domcontentloaded' });
    await page.waitForLoadState('networkidle').catch(() => {});

    const emailInput = await waitForAny(page, [
      'input[type="email"]',
      'input[name="loginfmt"]',
      'input[name="email"]',
      'input[type="text"][autocomplete="username"]'
    ], 15000);

    if (!emailInput) {
      throw new Error('Tidak menemukan form email login');
    }

    console.log(`Email dimasukkan: ${email}`);
    await emailInput.fill(email);
    await clickPrimary(page);

    const passwordInput = await waitForAny(page, [
      'input[type="password"]',
      'input[name="passwd"]',
      'input[name="password"]'
    ], 20000);

    if (!passwordInput) {
      throw new Error('Halaman password tidak muncul, kemungkinan akun memerlukan langkah tambahan');
    }

    await passwordInput.fill(password);
    await clickPrimary(page);

    console.log('Menunggu proses login dan MFA...');
    await page.waitForTimeout(5000);

    const mfaInput = await waitForAny(page, [
      'input[name="otc"]',
      'input[aria-label*="code" i]',
      'input[type="text"][autocomplete="one-time-code"]'
    ], 15000);

    let finalStatus = 'FAIL';

    if (mfaInput) {
      if (!mfaCode) {
        finalStatus = '2FA';
        await page.waitForTimeout(60000);
      } else {
        await mfaInput.fill(mfaCode);
        await clickPrimary(page);
      }
    }

    await page.waitForTimeout(5000);

    if (finalStatus === '2FA') {
      console.log(`[1]${email}   ${colorize('[2FA]', 'yellow')}`);
    } else {
      const title = await page.title();
      const url = page.url();
      const lowerTitle = (title || '').toLowerCase();
      const lowerUrl = (url || '').toLowerCase();

      const isAuthenticated =
        lowerUrl.includes('office.com') ||
        lowerUrl.includes('outlook.office.com') ||
        lowerUrl.includes('microsoft365.com') ||
        lowerUrl.includes('sharepoint.com') ||
        lowerUrl.includes('myaccount.microsoft.com') ||
        lowerUrl.includes('login.microsoftonline.com/common/oauth2') ||
        lowerTitle.includes('microsoft 365') ||
        lowerTitle.includes('outlook') ||
        lowerTitle.includes('your account') ||
        lowerTitle.includes('account') ||
        lowerUrl.includes('consent') ||
        lowerUrl.includes('sso') ||
        lowerUrl.includes('microsoftonline.com/common/SAS') ||
        lowerUrl.includes('microsoftonline.com/common/login') ||
        lowerUrl.includes('microsoftonline.com/common/');

      if (isAuthenticated) {
        console.log(`[1]${email}   ${colorize('[OK]', 'green')}`);
      } else {
        console.log(`[1]${email}   ${colorize('[FAIL]', 'red')}`);
      }
    }
  } catch (err) {
    console.error('Terjadi kesalahan:', err.message);
  } finally {
    if (browser) {
      await browser.close();
    }
  }
})();
