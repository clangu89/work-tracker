const fs = require('fs');
const path = require('path');

const ROOT_DIR = __dirname;
const DOCS_DIR = path.join(ROOT_DIR, 'docs');

const GAS_API_ENDPOINT = 'https://script.google.com/macros/s/AKfycbyox_uV2XJjgn96jK-XMKlr4vgucWG9FD1WqJR6Rhvu-L82El1nNhq7LyNvtwEBSjlrkg/exec';

function copyDirRecursive(src, dest) {
  if (!fs.existsSync(src)) return;
  if (!fs.existsSync(dest)) {
    fs.mkdirSync(dest, { recursive: true });
  }
  const entries = fs.readdirSync(src, { withFileTypes: true });
  for (const entry of entries) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyDirRecursive(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

function build() {
  console.log('🚀 Building Work Tracker + RDP Manager for GitHub Pages...');

  if (!fs.existsSync(DOCS_DIR)) {
    fs.mkdirSync(DOCS_DIR, { recursive: true });
  }

  const indexHtml = fs.readFileSync(path.join(ROOT_DIR, 'Index.html'), 'utf8');
  const stylesHtml = fs.readFileSync(path.join(ROOT_DIR, 'Styles.html'), 'utf8');
  const javascriptHtml = fs.readFileSync(path.join(ROOT_DIR, 'JavaScript.html'), 'utf8');
  const rdpViewHtml = fs.readFileSync(path.join(ROOT_DIR, 'RdpView.html'), 'utf8');
  const adminViewHtml = fs.readFileSync(path.join(ROOT_DIR, 'AdminView.html'), 'utf8');
  const rdpScriptHtml = fs.readFileSync(path.join(ROOT_DIR, 'RdpScript.html'), 'utf8');
  const adminScriptHtml = fs.readFileSync(path.join(ROOT_DIR, 'AdminScript.html'), 'utf8');
  const calendarViewHtml = fs.readFileSync(path.join(ROOT_DIR, 'CalendarView.html'), 'utf8');
  const calendarScriptHtml = fs.readFileSync(path.join(ROOT_DIR, 'CalendarScript.html'), 'utf8');

  // API Bridge Script to polyfill google.script.run for GitHub Pages
  const apiBridge = `
<script>
/**
 * =========================================================================
 * GOOGLE APPS SCRIPT API BRIDGE FOR GITHUB PAGES
 * Polyfill layer intercepting google.script.run and routing to Google Apps Script Web App
 * Includes Timeout, Cross-Origin Safeguards, Warmup Ping, and Exponential Backoff Retry
 * =========================================================================
 */
const GAS_API_ENDPOINT = '${GAS_API_ENDPOINT}';
window.GAS_API_ENDPOINT = GAS_API_ENDPOINT;

(function() {
  const isExternal = typeof window.google === 'undefined' || !window.google.script || !window.google.script.run;
  if (isExternal) {
    console.log('🌐 [API Bridge] Initialized on GitHub Pages / External Host');
    window.google = window.google || {};
    window.google.script = window.google.script || {};

    const RETRIABLE_ACTIONS = new Set([
      'loginWithPassword',
      'requestLoginPin',
      'verifyLoginPin',
      'getInitialData',
      'checkTasksUpdate',
      'getArchivedTasks',
      'getUsersList',
      'pingOnlineUser',
      'getOnlineUsers',
      'logoutOnlineUser',
      'getRdpInitialData',
      'getRdpGroups',
      'getRdpAuditLogs',
      'getAdminOverviewData',
      'getReportTasksData',
      'getSiteSchedules',
      'getCalendarBootstrapData',
      'getCalendarTriggerStatus',
      'testTelegramConnection'
    ]);

    function waitMs(ms) {
      return new Promise((resolve) => setTimeout(resolve, ms));
    }

    function isRetriableError(err, httpStatus) {
      if (httpStatus && (httpStatus === 429 || httpStatus >= 500)) return true;
      if (!err) return false;
      if (err.name === 'TypeError' || err.name === 'AbortError') return true;
      const msg = String(err.message || '');
      return /failed to fetch|networkerror|load failed|network request failed|timed? out|abort/i.test(msg);
    }

    function buildFriendlyError(err, propKey) {
      const rawMsg = String((err && err.message) || err || 'Unknown error');
      const isNetworkOrTimeout =
        (err && (err.name === 'TypeError' || err.name === 'AbortError')) ||
        /failed to fetch|networkerror|load failed|network request failed|timed? out/i.test(rawMsg);

      let displayMsg = rawMsg;
      if (typeof navigator !== 'undefined' && navigator.onLine === false) {
        displayMsg = 'ไม่พบการเชื่อมต่ออินเทอร์เน็ต กรุณาตรวจสอบ Wi-Fi หรือเครือข่ายแล้วลองใหม่อีกครั้ง';
      } else if (err && err.name === 'AbortError') {
        displayMsg = 'การเชื่อมต่อเซิร์ฟเวอร์ใช้เวลานานเกินกำหนด (Timeout) กรุณาลองใหม่อีกครั้ง';
      } else if (isNetworkOrTimeout) {
        displayMsg = 'ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้ชั่วคราว (Network / CORS Redirect) กรุณาลองกดใหม่อีกครั้ง';
      }

      const wrapped = new Error(displayMsg);
      wrapped.rawMessage = rawMsg;
      wrapped.isNetworkError = isNetworkOrTimeout;
      wrapped.action = propKey;
      return wrapped;
    }

    async function fetchWithTimeout(url, options, timeoutMs) {
      const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
      const timerId = controller ? setTimeout(() => controller.abort(), timeoutMs) : null;
      try {
        return await fetch(url, {
          ...options,
          credentials: 'omit',
          cache: 'no-store',
          redirect: 'follow',
          signal: controller ? controller.signal : undefined
        });
      } finally {
        if (timerId) clearTimeout(timerId);
      }
    }

    let activeWarmupPromise = null;
    window.warmupGasBackend = function(forceNew, timeoutMs) {
      if (!forceNew && activeWarmupPromise) return activeWarmupPromise;
      const pingTimeout = typeof timeoutMs === 'number' ? timeoutMs : 15000;
      const runPing = async () => {
        const startedAt = Date.now();
        try {
          const res = await fetchWithTimeout(
            GAS_API_ENDPOINT + '?action=ping&_t=' + Date.now(),
            { method: 'GET' },
            pingTimeout
          );
          if (!res.ok) return { ok: false, status: res.status, latencyMs: Date.now() - startedAt };
          const data = await res.json().catch(() => null);
          return { ok: Boolean(data && data.success), data: data, latencyMs: Date.now() - startedAt };
        } catch (err) {
          return { ok: false, error: err && err.message ? err.message : String(err), latencyMs: Date.now() - startedAt };
        } finally {
          setTimeout(() => { activeWarmupPromise = null; }, 2000);
        }
      };
      activeWarmupPromise = runPing();
      return activeWarmupPromise;
    };

    function createGasProxy(successCb, failureCb) {
      return new Proxy({}, {
        get(target, propKey) {
          if (propKey === 'withSuccessHandler') {
            return function(cb) { return createGasProxy(cb, failureCb); };
          }
          if (propKey === 'withFailureHandler') {
            return function(cb) { return createGasProxy(successCb, cb); };
          }
          return async function(...args) {
            const isUpload = String(propKey).toLowerCase().includes('upload');
            const timeoutMs = isUpload ? 60000 : 25000;
            const maxAttempts = RETRIABLE_ACTIONS.has(propKey) ? 3 : 1;

            let lastErr = null;
            for (let attempt = 1; attempt <= maxAttempts; attempt++) {
              let httpStatus = 0;
              try {
                const response = await fetchWithTimeout(GAS_API_ENDPOINT, {
                  method: 'POST',
                  headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                  body: JSON.stringify({ action: propKey, args: args })
                }, timeoutMs);

                httpStatus = response.status;
                if (!response.ok) {
                  throw new Error(\`HTTP \${response.status}: \${response.statusText}\`);
                }

                const data = await response.json();
                if (data && data.isServerException) {
                  throw new Error(data.message || data.error || 'Server Exception');
                }
                if (typeof successCb === 'function') {
                  successCb(data);
                }
                return data;
              } catch (err) {
                lastErr = err;
                const canRetry = attempt < maxAttempts && isRetriableError(err, httpStatus);
                console.warn(\`⚠️ [API Bridge] \${propKey} attempt \${attempt}/\${maxAttempts} failed:\`, err && err.message ? err.message : err);
                if (canRetry) {
                  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
                    break;
                  }
                  const backoffMs = 700 * Math.pow(2, attempt - 1);
                  await waitMs(backoffMs);
                  continue;
                }
                break;
              }
            }

            const finalErr = buildFriendlyError(lastErr, propKey);
            console.error(\`❌ [API Bridge Error] \${propKey}:\`, lastErr);
            if (typeof failureCb === 'function') {
              failureCb(finalErr);
            }
          };
        }
      });
    }

    window.google.script.run = createGasProxy();
  }
})();
</script>
`;

  // Inject styles, views, and scripts
  let output = indexHtml.replace(/<\?!=\s*include\('Styles'\);\s*\?>/g, stylesHtml);
  output = output.replace(/<\?!=\s*include\('RdpView'\);\s*\?>/g, rdpViewHtml);
  output = output.replace(/<\?!=\s*include\('AdminView'\);\s*\?>/g, adminViewHtml);
  output = output.replace(/<\?!=\s*include\('CalendarView'\);\s*\?>/g, calendarViewHtml);
  output = output.replace(/<\?!=\s*include\('JavaScript'\);\s*\?>/g, apiBridge + '\n' + javascriptHtml);
  output = output.replace(/<\?!=\s*include\('RdpScript'\);\s*\?>/g, rdpScriptHtml);
  output = output.replace(/<\?!=\s*include\('AdminScript'\);\s*\?>/g, adminScriptHtml);
  output = output.replace(/<\?!=\s*include\('CalendarScript'\);\s*\?>/g, calendarScriptHtml);

  // Fallback for state.webAppUrl in standalone build
  output = output.replace(
    /state\.webAppUrl = response\.webAppUrl \|\| '';/g,
    `state.webAppUrl = response.webAppUrl || '${GAS_API_ENDPOINT}';`
  );

  const targetFile = path.join(DOCS_DIR, 'index.html');
  fs.writeFileSync(targetFile, output, 'utf8');
  console.log(`✅ Generated standalone build at: ${targetFile}`);

  // Copy helper directory to docs/helper
  const helperSrc = path.join(ROOT_DIR, 'helper');
  const helperDest = path.join(DOCS_DIR, 'helper');
  if (fs.existsSync(helperSrc)) {
    copyDirRecursive(helperSrc, helperDest);
    console.log(`✅ Copied helper/ to docs/helper/`);
  }

  // Add .nojekyll so GitHub Pages doesn't ignore files
  fs.writeFileSync(path.join(DOCS_DIR, '.nojekyll'), '', 'utf8');
  console.log('✅ Created docs/.nojekyll');
}

build();
