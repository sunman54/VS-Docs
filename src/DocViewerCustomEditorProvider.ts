import * as vscode from 'vscode';
import * as path from 'path';

export class DocCustomDocument implements vscode.CustomDocument {
  constructor(public readonly uri: vscode.Uri) {}
  dispose(): void {}
}

export class DocViewerCustomEditorProvider implements vscode.CustomReadonlyEditorProvider<DocCustomDocument> {
  public static readonly viewType = 'vcdocs.docViewer';

  constructor(private readonly context: vscode.ExtensionContext) {}

  public static register(context: vscode.ExtensionContext): vscode.Disposable {
    const provider = new DocViewerCustomEditorProvider(context);
    return vscode.window.registerCustomEditorProvider(
      DocViewerCustomEditorProvider.viewType,
      provider,
      {
        webviewOptions: {
          retainContextWhenHidden: true,
        },
        supportsMultipleEditorsPerDocument: false,
      }
    );
  }

  async openCustomDocument(
    uri: vscode.Uri,
    _openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<DocCustomDocument> {
    return new DocCustomDocument(uri);
  }

  async resolveCustomEditor(
    document: DocCustomDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken
  ): Promise<void> {
    const webview = webviewPanel.webview;
    webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.joinPath(this.context.extensionUri, 'dist'),
      ],
    };

    webview.html = this.getHtmlForWebview(webview);

    const getThemeKindString = (kind: vscode.ColorThemeKind): 'dark' | 'light' | 'high-contrast' => {
      switch (kind) {
        case vscode.ColorThemeKind.Dark:
          return 'dark';
        case vscode.ColorThemeKind.HighContrast:
          return 'high-contrast';
        case vscode.ColorThemeKind.HighContrastLight:
        case vscode.ColorThemeKind.Light:
        default:
          return 'light';
      }
    };

    const sendDocumentData = async () => {
      try {
        const fileData = await vscode.workspace.fs.readFile(document.uri);
        const fileName = path.basename(document.uri.fsPath);
        const fileExt = path.extname(document.uri.fsPath).toLowerCase().replace('.', '');
        const currentTheme = getThemeKindString(vscode.window.activeColorTheme.kind);

        webview.postMessage({
          command: 'renderDocument',
          fileName,
          fileExt,
          buffer: fileData.buffer.slice(fileData.byteOffset, fileData.byteOffset + fileData.byteLength),
          theme: currentTheme,
        });
      } catch (err: any) {
        webview.postMessage({
          command: 'showError',
          message: `Dosya okunamadı: ${err?.message || String(err)}`,
        });
      }
    };

    // Listen to theme changes in VS Code
    const themeChangeDisposable = vscode.window.onDidChangeActiveColorTheme((theme) => {
      webview.postMessage({
        command: 'themeChanged',
        theme: getThemeKindString(theme.kind),
      });
    });

    // Handle incoming messages from Webview
    const messageDisposable = webview.onDidReceiveMessage(async (msg) => {
      switch (msg.command) {
        case 'webviewReady':
          await sendDocumentData();
          break;
        case 'copyToClipboard':
          await vscode.env.clipboard.writeText(msg.text);
          vscode.window.showInformationMessage(msg.infoMessage || 'Panoya kopyalandı!');
          break;
        case 'exportImages':
          await this.saveExtractedImages(document.uri, msg.images);
          break;
        case 'showErrorMessage':
          vscode.window.showErrorMessage(msg.text);
          break;
        case 'showInfoMessage':
          vscode.window.showInformationMessage(msg.text);
          break;
      }
    });

    webviewPanel.onDidDispose(() => {
      themeChangeDisposable.dispose();
      messageDisposable.dispose();
    });
  }

  private async saveExtractedImages(docUri: vscode.Uri, images: Array<{ name: string; base64: string }>): Promise<void> {
    if (!images || images.length === 0) {
      vscode.window.showInformationMessage('Dokümanda dışa aktarılacak görsel bulunamadı.');
      return;
    }

    try {
      const docDir = path.dirname(docUri.fsPath);
      const docName = path.basename(docUri.fsPath, path.extname(docUri.fsPath));
      const targetDir = path.join(docDir, `${docName}_gorseller`);
      const targetUri = vscode.Uri.file(targetDir);

      await vscode.workspace.fs.createDirectory(targetUri);

      for (const img of images) {
        const fileUri = vscode.Uri.file(path.join(targetDir, img.name));
        const buffer = Buffer.from(img.base64, 'base64');
        await vscode.workspace.fs.writeFile(fileUri, buffer);
      }

      vscode.window.showInformationMessage(
        `${images.length} adet görsel başarıyla kaydedildi:\n${targetDir}`
      );
    } catch (err: any) {
      vscode.window.showErrorMessage(`Görseller kaydedilemedi: ${err?.message || String(err)}`);
    }
  }

  private getHtmlForWebview(webview: vscode.Webview): string {
    const distUri = (file: string) =>
      webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'dist', file));

    const scriptUri = distUri('webview.js');
    const styleUri = distUri('webview.css');
    const nonce = getNonce();

    return `<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="
    default-src 'none';
    style-src ${webview.cspSource} 'unsafe-inline';
    script-src 'nonce-${nonce}';
    img-src ${webview.cspSource} data: blob:;
    font-src ${webview.cspSource};
  ">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="stylesheet" href="${styleUri}">
  <title>VS Docs</title>
</head>
<body class="loading">
  <!-- Toolbar -->
  <header id="app-toolbar" class="app-toolbar">
    <div class="toolbar-left">
      <span id="doc-title" class="doc-title" title="Doküman">Yükleniyor...</span>
      <span id="doc-badge" class="doc-badge">DOC</span>
      <button id="btn-toggle-outline" class="toolbar-btn text-btn hidden" title="İçindekiler / Başlık Ağacı">
        <span>📑</span>
        <span>İçindekiler</span>
      </button>
    </div>

    <div class="toolbar-center">
      <!-- Sheet Selector for Excel -->
      <div id="sheet-selector-container" class="toolbar-group hidden">
        <label for="sheet-select" class="toolbar-label">Sayfa:</label>
        <select id="sheet-select" class="toolbar-select"></select>
      </div>

      <!-- Delimiter Selector for CSV -->
      <div id="delimiter-selector-container" class="toolbar-group hidden">
        <label for="delimiter-select" class="toolbar-label">Ayraç:</label>
        <select id="delimiter-select" class="toolbar-select">
          <option value="auto">Otomatik</option>
          <option value=",">Virgül (,)</option>
          <option value=";">Noktalı Virgül (;)</option>
          <option value="&#9;">Tab (\t)</option>
          <option value="|">Dikey Çizgi (|)</option>
        </select>
      </div>

      <!-- Zoom Controls for Word -->
      <div id="zoom-controls" class="toolbar-group">
        <button id="btn-zoom-out" class="toolbar-btn" title="Uzaklaştır (−)">−</button>
        <span id="zoom-level" class="zoom-level">100%</span>
        <button id="btn-zoom-in" class="toolbar-btn" title="Yakınlaştır (+)">+</button>
        <button id="btn-zoom-reset" class="toolbar-btn text-btn" title="Sıfırla">1:1</button>
      </div>

      <!-- Search Box -->
      <div id="search-box-container" class="toolbar-group search-group">
        <input type="text" id="search-input" class="search-input" placeholder="Dokümanda ara..." />
        <span id="search-count" class="search-count"></span>
      </div>
    </div>

    <div class="toolbar-right">
      <!-- Developer Tools: Export/Copy -->
      <div id="dev-export-group" class="toolbar-group">
        <button id="btn-copy-md" class="toolbar-btn text-btn" title="Tabloyu Markdown formatında kopyala">
          <span>📋</span>
          <span>Markdown</span>
        </button>
        <button id="btn-copy-json" class="toolbar-btn text-btn" title="Tabloyu JSON dizisi olarak kopyala">
          <span>{ }</span>
          <span>JSON</span>
        </button>
        <button id="btn-export-images" class="toolbar-btn text-btn hidden" title="Dokümandaki gömülü resimleri dışa aktar">
          <span>🖼️</span>
          <span>Görseller</span>
        </button>
      </div>

      <!-- Print Button -->
      <button id="btn-print" class="toolbar-btn text-btn" title="Yazdır veya PDF olarak kaydet">
        <span>🖨️</span>
        <span>Yazdır</span>
      </button>

      <!-- Theme Switcher -->
      <button id="btn-theme-toggle" class="toolbar-btn text-btn" title="Tema Değiştir (Açık/Koyu/Sistem)">
        <span id="theme-icon">🌓</span>
        <span id="theme-text">Otomatik</span>
      </button>
    </div>
  </header>

  <!-- Error Alert Container -->
  <div id="error-banner" class="error-banner hidden">
    <span id="error-message"></span>
    <button id="btn-close-error" class="close-error-btn">✕</button>
  </div>

  <!-- Loading Spinner -->
  <div id="loading-overlay" class="loading-overlay">
    <div class="spinner"></div>
    <p id="loading-text">Doküman hazırlanıyor...</p>
  </div>

  <!-- Main Content Layout (Sidebar + Viewer) -->
  <div class="app-body-layout">
    <aside id="outline-sidebar" class="outline-sidebar hidden">
      <div class="outline-header">
        <span>İçindekiler</span>
        <button id="btn-close-outline" class="close-outline-btn">✕</button>
      </div>
      <div class="outline-list"></div>
    </aside>

    <!-- Main Viewer Content -->
    <main id="viewer-container" class="viewer-container">
      <div id="document-content" class="document-content"></div>
    </main>
  </div>

  <!-- Bottom Status Bar -->
  <footer id="doc-status-bar" class="doc-status-bar">
    <span id="status-selection" class="status-item">Hazır</span>
    <span id="status-stats" class="status-item"></span>
  </footer>

  <script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
  }
}

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

