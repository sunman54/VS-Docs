const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

const isWatch = process.argv.includes('--watch');

async function build() {
  if (!fs.existsSync('dist')) {
    fs.mkdirSync('dist', { recursive: true });
  }

  // Copy CSS to dist
  const cssSource = path.join(__dirname, 'src', 'webview', 'styles.css');
  const cssDest = path.join(__dirname, 'dist', 'webview.css');
  if (fs.existsSync(cssSource)) {
    fs.copyFileSync(cssSource, cssDest);
  }

  // 1. Extension Host Bundle
  const extensionCtx = await esbuild.context({
    entryPoints: ['src/extension.ts'],
    bundle: true,
    format: 'cjs',
    platform: 'node',
    outfile: 'dist/extension.js',
    external: ['vscode'],
    sourcemap: true,
    minify: !isWatch,
    keepNames: true,
    logLevel: 'info',
  });

  // 2. Webview Client Bundle (includes docx-preview, xlsx, dompurify)
  const webviewCtx = await esbuild.context({
    entryPoints: ['src/webview/main.ts'],
    bundle: true,
    format: 'iife',
    platform: 'browser',
    outfile: 'dist/webview.js',
    sourcemap: true,
    minify: !isWatch,
    keepNames: true,
    logLevel: 'info',
    // Node polyfills or stubs if needed by xlsx in browser
    define: {
      'process.env.NODE_ENV': isWatch ? '"development"' : '"production"',
      'global': 'window',
    },
  });

  if (isWatch) {
    console.log('Watching for changes...');
    await Promise.all([extensionCtx.watch(), webviewCtx.watch()]);
  } else {
    await Promise.all([extensionCtx.rebuild(), webviewCtx.rebuild()]);
    await Promise.all([extensionCtx.dispose(), webviewCtx.dispose()]);
    console.log('Build completed successfully.');
  }
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
