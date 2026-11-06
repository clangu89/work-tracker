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
 * =========================================================================
 */
const GAS_API_ENDPOINT = '${GAS_API_ENDPOINT}';

(function() {
  const isExternal = typeof window.google === 'undefined' || !window.google.script || !window.google.script.run;
  if (isExternal) {
    console.log('🌐 [API Bridge] Initialized on GitHub Pages / External Host');
    window.google = window.google || {};
    window.google.script = window.google.script || {};

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
            try {
              const response = await fetch(GAS_API_ENDPOINT, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: propKey, args: args })
              });
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
              console.error(\`❌ [API Bridge Error] \${propKey}:\`, err);
              if (typeof failureCb === 'function') {
                failureCb(err);
              }
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
