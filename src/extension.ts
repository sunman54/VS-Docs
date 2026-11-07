import * as vscode from 'vscode';
import { DocViewerCustomEditorProvider } from './DocViewerCustomEditorProvider';

export function activate(context: vscode.ExtensionContext) {
  // Register the custom read-only editor provider
  const providerDisposable = DocViewerCustomEditorProvider.register(context);
  context.subscriptions.push(providerDisposable);

  // Command to open active or selected document with our viewer
  const openViewerCommand = vscode.commands.registerCommand('vcdocs.openViewer', async (uri?: vscode.Uri) => {
    let targetUri = uri;
    if (!targetUri) {
      targetUri = vscode.window.activeTextEditor?.document.uri;
    }
    if (!targetUri) {
      vscode.window.showWarningMessage('Lütfen görüntülemek için bir dosya seçin.');
      return;
    }

    await vscode.commands.executeCommand('vscode.openWith', targetUri, DocViewerCustomEditorProvider.viewType);
  });

  context.subscriptions.push(openViewerCommand);
}

export function deactivate() {}

