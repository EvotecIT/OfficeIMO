import * as vscode from 'vscode';
import { spawn } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';

const allowedCliArtifactNames = new Set([
  'officeimo.markup.cli',
  'officeimo.markup.cli.csproj',
  'officeimo.markup.cli.dll',
  'officeimo.markup.cli.exe'
]);

let warnedAboutInvalidCliPath = false;

type CliResult = {
  stdout: string;
  stderr: string;
  code: number | null;
};

type ProcessResult = CliResult;

type MarkupEnvelope = {
  Document?: MarkupDocument;
  Diagnostics?: MarkupDiagnostic[];
};

type ValidationEnvelope = {
  Diagnostics?: MarkupDiagnostic[];
  HasErrors?: boolean;
};

type MarkupDocument = {
  Profile?: string;
  Metadata?: Record<string, string>;
  Blocks?: MarkupBlock[];
};

type MarkupPlacement = {
  X?: string;
  Y?: string;
  Width?: string;
  Height?: string;
};

type MarkupStyle = {
  Name?: string;
  FontName?: string;
  FontSize?: number;
  Bold?: boolean;
  Italic?: boolean;
  TextColor?: string;
  FillColor?: string;
  BorderColor?: string;
  TextAlign?: string;
};

type MarkupTransitionDetails = {
  RawText?: string;
  Effect?: string;
  ResolvedIdentifier?: string;
  Attributes?: Record<string, string>;
};

type MarkupBlock = {
  Kind?: string;
  SourceText?: string;
  Text?: string;
  Title?: string;
  Level?: number;
  Ordered?: boolean;
  Start?: number;
  Layout?: string;
  Section?: string;
  Transition?: string;
  TransitionDetails?: MarkupTransitionDetails;
  Background?: string;
  Notes?: string;
  Placement?: string;
  Columns?: number;
  Blocks?: MarkupBlock[];
  Items?: Array<{ Text?: string }>;
  Language?: string;
  Content?: string;
  Source?: string;
  Alt?: string;
  Width?: number;
  Height?: number;
  RenderAsImage?: boolean;
  Headers?: string[];
  ChartType?: string;
  Rows?: string[][];
  Name?: string;
  PageSize?: string;
  Orientation?: string;
  MinLevel?: number;
  MaxLevel?: number;
  Address?: string;
  Sheet?: string;
  Cell?: string;
  Expression?: string;
  Range?: string;
  HasHeader?: boolean;
  Target?: string;
  Style?: string;
  NumberFormat?: string;
  Gap?: string;
  ColumnKind?: string;
  WidthText?: string;
  Position?: MarkupPlacement;
  ResolvedStyle?: MarkupStyle;
  Command?: string;
  Body?: string;
  Markdown?: string;
  Attributes?: Record<string, string>;
};

type SlidePreviewBackground = {
  className: string;
  style: string;
  overlayHtml: string;
};

type PreviewTheme = {
  background: string;
  surface: string;
  panel: string;
  panelBorder: string;
  text: string;
  textSecondary: string;
  textMuted: string;
  accent: string;
  accentDark: string;
  accentLight: string;
  accent2: string;
  accent3: string;
  warning: string;
};

type MarkupDiagnostic = {
  Severity?: string;
  Message?: string;
  NodeKind?: string;
  NodeSourceText?: string;
};

type PreviewRenderMode = 'document' | 'slide' | 'workbook';
type ChartPreviewRow = { category: string; values: number[] };
type ExportTarget = 'pptx' | 'xlsx' | 'docx';
type WorkbookPreviewCellStyle = {
  numberFormat?: string;
  fillColor?: string;
  textColor?: string;
  borderColor?: string;
  borderStyle?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  textAlign?: string;
  verticalAlign?: string;
  wrap?: boolean;
};
type WorkbookPreviewCell = {
  text?: string;
  formula?: string;
  style: WorkbookPreviewCellStyle;
};
type WorkbookPreviewTable = {
  name?: string;
  hasHeader: boolean;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
};
type WorkbookPreviewSheet = {
  name: string;
  cells: Map<string, WorkbookPreviewCell>;
  tables: WorkbookPreviewTable[];
  charts: MarkupBlock[];
  extras: MarkupBlock[];
};

let diagnostics: vscode.DiagnosticCollection;
const validationTimers = new Map<string, ReturnType<typeof setTimeout>>();
const previewTimers = new Map<string, ReturnType<typeof setTimeout>>();
const previewPanels = new Map<string, vscode.WebviewPanel>();
const previewVersions = new Map<string, number>();

export function activate(context: vscode.ExtensionContext): void {
  diagnostics = vscode.languages.createDiagnosticCollection('officeimo-markup');
  context.subscriptions.push(diagnostics);

  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.preview', (resource?: vscode.Uri) => previewActiveDocument(context, resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.validate', (resource?: vscode.Uri) => validateActiveDocument(context, true, resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.emitCSharp', (resource?: vscode.Uri) => emitActiveDocument(context, 'csharp', resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.emitPowerShell', (resource?: vscode.Uri) => emitActiveDocument(context, 'powershell', resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.emitCSharpToFile', (resource?: vscode.Uri) => emitDocumentToFile(context, 'csharp', resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.emitPowerShellToFile', (resource?: vscode.Uri) => emitDocumentToFile(context, 'powershell', resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.generateArtifacts', (resource?: vscode.Uri) => generateArtifacts(context, resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.exportOffice', (resource?: vscode.Uri) => exportOffice(context, resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.exportOfficeAndOpen', (resource?: vscode.Uri) => exportOfficeAndOpen(context, resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.exportPowerPoint', (resource?: vscode.Uri) => exportPowerPoint(context, resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.exportExcel', (resource?: vscode.Uri) => exportExcel(context, resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.exportWord', (resource?: vscode.Uri) => exportWord(context, resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.openOutputFolder', (resource?: vscode.Uri) => openOutputFolder(resource)));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.openGeneratedCSharp', (resource?: vscode.Uri) => openGeneratedCodeFile(resource, 'csharp')));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.openGeneratedPowerShell', (resource?: vscode.Uri) => openGeneratedCodeFile(resource, 'powershell')));
  context.subscriptions.push(vscode.commands.registerCommand('officeimoMarkup.installMermaidRenderer', () => installMermaidRenderer(context)));

  context.subscriptions.push(vscode.workspace.onDidOpenTextDocument((document) => scheduleValidation(context, document)));
  context.subscriptions.push(vscode.workspace.onDidSaveTextDocument((document) => scheduleValidation(context, document, 0)));
  context.subscriptions.push(vscode.workspace.onDidChangeTextDocument((event) => {
    scheduleValidation(context, event.document);
    schedulePreviewRefresh(context, event.document);
  }));
  context.subscriptions.push(vscode.workspace.onDidSaveTextDocument((document) => schedulePreviewRefresh(context, document, 0)));

  for (const document of vscode.workspace.textDocuments) {
    scheduleValidation(context, document, 0);
  }
}

export function deactivate(): void {
  diagnostics?.dispose();
  previewTimers.forEach((timer) => clearTimeout(timer));
  validationTimers.forEach((timer) => clearTimeout(timer));
  previewTimers.clear();
  validationTimers.clear();
  previewPanels.clear();
}

async function previewActiveDocument(context: vscode.ExtensionContext, resource?: vscode.Uri): Promise<void> {
  const document = await activePreviewDocument(resource);
  if (!document) {
    return;
  }

  if (!isOfficeMarkupContent(document)) {
    await vscode.commands.executeCommand('markdown.showPreviewToSide', document.uri);
    return;
  }

  const key = document.uri.toString();
  let panel = previewPanels.get(key);
  if (!panel) {
    panel = vscode.window.createWebviewPanel(
      'officeimoMarkupPreview',
      `OfficeIMO Preview: ${path.basename(document.fileName)}`,
      vscode.ViewColumn.Beside,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
        localResourceRoots: previewLocalResourceRoots(context, document)
      }
    );
    panel.onDidDispose(() => {
      previewPanels.delete(key);
      previewVersions.delete(key);
      const timer = previewTimers.get(key);
      if (timer) {
        clearTimeout(timer);
        previewTimers.delete(key);
      }
    });
    panel.webview.onDidReceiveMessage(async (message) => {
      const resource = vscode.Uri.parse(key);
      try {
        switch (message?.command) {
          case 'refresh': {
            const previewDocument = await vscode.workspace.openTextDocument(resource);
            await updatePreviewPanel(context, previewDocument, panel!, true);
            break;
          }
          case 'validate':
            await vscode.commands.executeCommand('officeimoMarkup.validate', resource);
            break;
          case 'exportOfficeAndOpen':
            await vscode.commands.executeCommand('officeimoMarkup.exportOfficeAndOpen', resource);
            break;
          case 'generateArtifacts':
            await vscode.commands.executeCommand('officeimoMarkup.generateArtifacts', resource);
            break;
          case 'openOutputFolder':
            await vscode.commands.executeCommand('officeimoMarkup.openOutputFolder', resource);
            break;
        }
      } catch (error) {
        void vscode.window.showErrorMessage(`OfficeIMO Markup preview action failed: ${String(error)}`);
      }
    });
    previewPanels.set(key, panel);
  } else {
    panel.reveal(vscode.ViewColumn.Beside, true);
  }

  await updatePreviewPanel(context, document, panel, true);
}

async function validateActiveDocument(context: vscode.ExtensionContext, showMessage: boolean, resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  await validateDocument(context, document, showMessage);
}

async function emitActiveDocument(context: vscode.ExtensionContext, target: 'csharp' | 'powershell', resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const outputText = await emitTargetContent(context, document, target);
  if (!outputText) {
    return;
  }

  const output = await vscode.workspace.openTextDocument({
    content: outputText,
    language: target === 'csharp' ? 'csharp' : 'powershell'
  });
  await vscode.window.showTextDocument(output, vscode.ViewColumn.Beside);
}

async function emitDocumentToFile(context: vscode.ExtensionContext, target: 'csharp' | 'powershell', resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const outputText = await emitTargetContent(context, document, target);
  if (!outputText) {
    return;
  }

  const output = await pickCodegenPath(document, target);
  if (!output) {
    return;
  }

  await vscode.workspace.fs.writeFile(output, new TextEncoder().encode(outputText));
  await showArtifactSaved(output, target === 'csharp' ? 'C# file' : 'PowerShell file');
}

async function generateArtifacts(context: vscode.ExtensionContext, resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const profileTarget = await resolveExportTarget(context, document, 'Choose the Office output for generated artifacts');
  if (!profileTarget) {
    return;
  }

  const csharp = await emitTargetContent(context, document, 'csharp');
  if (!csharp) {
    return;
  }

  const powershell = await emitTargetContent(context, document, 'powershell');
  if (!powershell) {
    return;
  }

  const outputDirectory = await resolveOutputDirectory(document);
  const csharpPath = await resolveGeneratedCodePath(document, 'csharp');
  const powershellPath = await resolveGeneratedCodePath(document, 'powershell');
  const officePath = vscode.Uri.file(path.join(outputDirectory.fsPath, `${path.basename(document.fileName, path.extname(document.fileName))}.${profileTarget}`));

  await vscode.workspace.fs.writeFile(csharpPath, new TextEncoder().encode(csharp));
  await vscode.workspace.fs.writeFile(powershellPath, new TextEncoder().encode(powershell));

  const saveLabel = profileTarget === 'pptx'
    ? 'Export PowerPoint'
    : profileTarget === 'docx'
      ? 'Export Word Document'
      : 'Export Excel Workbook';

  const exportResult = await runExportToPath(context, document, profileTarget, officePath, saveLabel);
  if (!exportResult) {
    return;
  }

  const action = await vscode.window.showInformationMessage(
    `Generated C#, PowerShell, and ${describeExportTarget(profileTarget)} artifacts in ${outputDirectory.fsPath}.`,
    'Open Office File',
    'Open C#',
    'Open PowerShell',
    'Reveal Folder'
  );

  if (action === 'Open Office File') {
    await vscode.env.openExternal(officePath);
  } else if (action === 'Open C#') {
    await openUriInEditor(csharpPath);
  } else if (action === 'Open PowerShell') {
    await openUriInEditor(powershellPath);
  } else if (action === 'Reveal Folder') {
    await vscode.commands.executeCommand('revealFileInOS', officePath);
  }
}

async function exportPowerPoint(context: vscode.ExtensionContext, resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const output = await pickExportPath(document, 'pptx', 'Export PowerPoint');
  if (!output) {
    return;
  }

  const exportArgs = ['--target', 'pptx', '--output', output.fsPath];
  const config = vscode.workspace.getConfiguration('officeimoMarkup');
  if (!config.get<boolean>('renderMermaidOnExport', true)) {
    exportArgs.push('--no-mermaid');
  }

  const mermaidCliPath = config.get<string>('mermaidCliPath', '').trim();
  const localMermaidRenderer = mermaidCliPath || findLocalMermaidRenderer(context);
  if (localMermaidRenderer) {
    exportArgs.push('--mermaid-renderer', localMermaidRenderer);
  }

  const result = await runCli(context, document, 'export', exportArgs);
  if (result.code !== 0) {
    vscode.window.showErrorMessage(result.stderr || 'OfficeIMO Markup PowerPoint export failed.');
    return;
  }

  await showExportSuccess(output, 'PowerPoint');
}

async function exportExcel(context: vscode.ExtensionContext, resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const output = await pickExportPath(document, 'xlsx', 'Export Excel Workbook');
  if (!output) {
    return;
  }

  const result = await runCli(context, document, 'export', ['--target', 'xlsx', '--output', output.fsPath]);
  if (result.code !== 0) {
    vscode.window.showErrorMessage(result.stderr || 'OfficeIMO Markup Excel export failed.');
    return;
  }

  await showExportSuccess(output, 'Excel workbook');
}

async function exportWord(context: vscode.ExtensionContext, resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const output = await pickExportPath(document, 'docx', 'Export Word Document');
  if (!output) {
    return;
  }

  const result = await runCli(context, document, 'export', ['--target', 'docx', '--output', output.fsPath]);
  if (result.code !== 0) {
    vscode.window.showErrorMessage(result.stderr || 'OfficeIMO Markup Word export failed.');
    return;
  }

  await showExportSuccess(output, 'Word document');
}

async function exportOffice(context: vscode.ExtensionContext, resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const target = await resolveExportTarget(context, document, 'Choose the Office output for this markup file');
  if (!target) {
    return;
  }

  if (target === 'pptx') {
    await exportPowerPoint(context, document.uri);
  } else if (target === 'docx') {
    await exportWord(context, document.uri);
  } else {
    await exportExcel(context, document.uri);
  }
}

async function exportOfficeAndOpen(context: vscode.ExtensionContext, resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const target = await resolveExportTarget(context, document, 'Choose the Office output to export and open');
  if (!target) {
    return;
  }

  const saveLabel = target === 'pptx'
    ? 'Export PowerPoint'
    : target === 'docx'
      ? 'Export Word Document'
      : 'Export Excel Workbook';
  const output = await exportDocumentTarget(context, document, target, saveLabel);
  if (output) {
    await vscode.env.openExternal(output);
  }
}

async function openOutputFolder(resource?: vscode.Uri): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const outputDirectory = await resolveOutputDirectory(document);
  await vscode.commands.executeCommand('revealFileInOS', outputDirectory);
}

async function openGeneratedCodeFile(resource: vscode.Uri | undefined, target: 'csharp' | 'powershell'): Promise<void> {
  const document = await activeMarkupDocument(resource);
  if (!document) {
    return;
  }

  const output = await resolveGeneratedCodePath(document, target);
  if (!await uriExists(output)) {
    vscode.window.showInformationMessage(
      `${target === 'csharp' ? 'Generated C#' : 'Generated PowerShell'} file does not exist yet. Run Generate Artifacts or Generate ${target === 'csharp' ? 'C# File' : 'PowerShell File'} first.`
    );
    return;
  }

  await openUriInEditor(output);
}

async function installMermaidRenderer(context: vscode.ExtensionContext): Promise<void> {
  const existing = findLocalMermaidRenderer(context);
  if (existing) {
    await rememberMermaidRenderer(existing);
    vscode.window.showInformationMessage(`OfficeIMO Markup Mermaid renderer is ready: ${existing}`);
    return;
  }

  const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  const installRoot = localMermaidInstallRoot(context);
  await fs.promises.mkdir(installRoot, { recursive: true });

  const result = await vscode.window.withProgress({
    location: vscode.ProgressLocation.Notification,
    cancellable: false,
    title: 'Installing OfficeIMO Markup Mermaid renderer'
  }, async (progress) => {
    progress.report({ message: 'Installing @mermaid-js/mermaid-cli locally for this VS Code profile...' });
    return runProcess(npm, ['install', '--prefix', installRoot, '@mermaid-js/mermaid-cli'], installRoot);
  });

  if (result.code !== 0) {
    vscode.window.showErrorMessage(result.stderr || result.stdout || 'Mermaid renderer installation failed.');
    return;
  }

  const renderer = findLocalMermaidRenderer(context);
  if (!renderer) {
    vscode.window.showWarningMessage('Mermaid CLI installed, but mmdc was not found in the expected local tools folder.');
    return;
  }

  await rememberMermaidRenderer(renderer);
  vscode.window.showInformationMessage(`OfficeIMO Markup Mermaid renderer installed: ${renderer}`);
}

async function exportDocumentTarget(context: vscode.ExtensionContext, document: vscode.TextDocument, target: ExportTarget, saveLabel: string): Promise<vscode.Uri | undefined> {
  const output = await pickExportPath(document, target, saveLabel);
  if (!output) {
    return undefined;
  }

  return runExportToPath(context, document, target, output, saveLabel);
}

async function runExportToPath(
  context: vscode.ExtensionContext,
  document: vscode.TextDocument,
  target: ExportTarget,
  output: vscode.Uri,
  saveLabel: string
): Promise<vscode.Uri | undefined> {

  const exportArgs = ['--target', target, '--output', output.fsPath];
  if (target === 'pptx') {
    const config = vscode.workspace.getConfiguration('officeimoMarkup');
    if (!config.get<boolean>('renderMermaidOnExport', true)) {
      exportArgs.push('--no-mermaid');
    }

    const mermaidCliPath = config.get<string>('mermaidCliPath', '').trim();
    const localMermaidRenderer = mermaidCliPath || findLocalMermaidRenderer(context);
    if (localMermaidRenderer) {
      exportArgs.push('--mermaid-renderer', localMermaidRenderer);
    }
  }

  const result = await runCli(context, document, 'export', exportArgs);
  if (result.code !== 0) {
    vscode.window.showErrorMessage(result.stderr || `OfficeIMO Markup ${saveLabel.toLowerCase()} failed.`);
    return undefined;
  }

  return output;
}

async function resolveExportTarget(
  context: vscode.ExtensionContext,
  document: vscode.TextDocument,
  placeHolder: string
): Promise<ExportTarget | undefined> {
  const profile = await resolveDocumentProfile(context, document);
  switch (profile) {
    case 'presentation':
      return 'pptx';
    case 'document':
      return 'docx';
    case 'workbook':
      return 'xlsx';
    default: {
      const picked = await vscode.window.showQuickPick([
        { label: 'PowerPoint presentation', target: 'pptx' as ExportTarget },
        { label: 'Word document', target: 'docx' as ExportTarget },
        { label: 'Excel workbook', target: 'xlsx' as ExportTarget }
      ], { placeHolder });
      return picked?.target;
    }
  }
}

async function emitTargetContent(
  context: vscode.ExtensionContext,
  document: vscode.TextDocument,
  target: 'csharp' | 'powershell'
): Promise<string | undefined> {
  const result = await runCli(context, document, 'emit', ['--target', target]);
  if (result.code !== 0) {
    vscode.window.showErrorMessage(result.stderr || `OfficeIMO Markup emit failed for ${target}.`);
    return undefined;
  }

  return result.stdout;
}

async function pickExportPath(document: vscode.TextDocument, target: ExportTarget, saveLabel: string): Promise<vscode.Uri | undefined> {
  const extension = target;
  const outputDirectory = await resolveOutputDirectory(document);
  const defaultUri = vscode.Uri.file(path.join(
    outputDirectory.fsPath,
    `${path.basename(document.fileName, path.extname(document.fileName))}.${extension}`
  ));

  const filters = target === 'pptx'
    ? { PowerPoint: ['pptx'] }
    : target === 'docx'
      ? { Word: ['docx'] }
      : { Excel: ['xlsx'] };

  return vscode.window.showSaveDialog({
    defaultUri,
    filters,
    saveLabel
  });
}

async function pickCodegenPath(
  document: vscode.TextDocument,
  target: 'csharp' | 'powershell'
): Promise<vscode.Uri | undefined> {
  const extension = target === 'csharp' ? 'cs' : 'ps1';
  const saveLabel = target === 'csharp' ? 'Generate C# File' : 'Generate PowerShell File';
  const defaultUri = await resolveGeneratedCodePath(document, target);

  const filters = target === 'csharp'
    ? { 'C#': ['cs'] }
    : { 'PowerShell': ['ps1'] };

  return vscode.window.showSaveDialog({
    defaultUri,
    filters,
    saveLabel
  });
}

async function showExportSuccess(output: vscode.Uri, kind: string): Promise<void> {
  await showArtifactSaved(output, kind);
}

async function showArtifactSaved(output: vscode.Uri, kind: string): Promise<void> {
  const action = await vscode.window.showInformationMessage(
    `${kind} saved to ${output.fsPath}.`,
    'Open',
    'Reveal'
  );

  if (action === 'Open') {
    await vscode.env.openExternal(output);
  } else if (action === 'Reveal') {
    await vscode.commands.executeCommand('revealFileInOS', output);
  }
}

async function resolveGeneratedCodePath(
  document: vscode.TextDocument,
  target: 'csharp' | 'powershell'
): Promise<vscode.Uri> {
  const extension = target === 'csharp' ? 'cs' : 'ps1';
  const outputDirectory = await resolveOutputDirectory(document);
  return vscode.Uri.file(path.join(
    outputDirectory.fsPath,
    `${path.basename(document.fileName, path.extname(document.fileName))}.generated.${extension}`
  ));
}

function describeExportTarget(target: ExportTarget): string {
  switch (target) {
    case 'pptx':
      return 'PowerPoint';
    case 'docx':
      return 'Word';
    case 'xlsx':
      return 'Excel';
  }
}

async function resolveOutputDirectory(document: vscode.TextDocument): Promise<vscode.Uri> {
  const configuration = vscode.workspace.getConfiguration('officeimoMarkup', document.uri);
  const mode = configuration.get<string>('outputDirectoryMode', 'sourceDirectory');
  const sourceDirectory = path.dirname(document.fileName);

  if (mode !== 'generatedSubfolder') {
    return vscode.Uri.file(sourceDirectory);
  }

  const configuredName = configuration.get<string>('outputSubfolderName', 'generated').trim();
  const subfolderName = configuredName || 'generated';
  const outputDirectory = vscode.Uri.file(path.join(sourceDirectory, subfolderName));
  await vscode.workspace.fs.createDirectory(outputDirectory);
  return outputDirectory;
}

async function uriExists(uri: vscode.Uri): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(uri);
    return true;
  } catch {
    return false;
  }
}

async function openUriInEditor(uri: vscode.Uri): Promise<void> {
  const opened = await vscode.workspace.openTextDocument(uri);
  await vscode.window.showTextDocument(opened, vscode.ViewColumn.Beside);
}

async function activePreviewDocument(resource?: vscode.Uri): Promise<vscode.TextDocument | undefined> {
  const document = resource ? await vscode.workspace.openTextDocument(resource) : vscode.window.activeTextEditor?.document;
  if (!document || !isMarkupCommandCandidate(document)) {
    vscode.window.showInformationMessage('Open a Markdown, .office.md, or .omd file first.');
    return undefined;
  }

  return document;
}

async function activeMarkupDocument(resource?: vscode.Uri): Promise<vscode.TextDocument | undefined> {
  const document = await activePreviewDocument(resource);
  if (!document) {
    return undefined;
  }

  if (!isOfficeMarkupContent(document)) {
    vscode.window.showInformationMessage('Open an OfficeIMO markup file (.omd, .office.md, or Markdown with OfficeIMO profile/directives) first.');
    return undefined;
  }

  return document;
}

function scheduleValidation(context: vscode.ExtensionContext, document: vscode.TextDocument, delay?: number): void {
  if (!isMarkupDocument(document)) {
    return;
  }

  const key = document.uri.toString();
  const existing = validationTimers.get(key);
  if (existing) {
    clearTimeout(existing);
  }

  const configuredDelay = vscode.workspace.getConfiguration('officeimoMarkup').get<number>('validateDebounceMs', 650);
  const timer = setTimeout(() => {
    validationTimers.delete(key);
    validateDocument(context, document, false).catch((error) => {
      diagnostics.set(document.uri, [new vscode.Diagnostic(new vscode.Range(0, 0, 0, 1), String(error), vscode.DiagnosticSeverity.Warning)]);
    });
  }, delay ?? configuredDelay);

  validationTimers.set(key, timer);
}

function schedulePreviewRefresh(context: vscode.ExtensionContext, document: vscode.TextDocument, delay?: number): void {
  if (!isOfficeMarkupContent(document)) {
    return;
  }

  const config = vscode.workspace.getConfiguration('officeimoMarkup');
  if (!config.get<boolean>('previewAutoRefresh', true)) {
    return;
  }

  const key = document.uri.toString();
  const panel = previewPanels.get(key);
  if (!panel) {
    return;
  }

  const existing = previewTimers.get(key);
  if (existing) {
    clearTimeout(existing);
  }

  const configuredDelay = config.get<number>('previewDebounceMs', 350);
  const timer = setTimeout(() => {
    previewTimers.delete(key);
    updatePreviewPanel(context, document, panel, false).catch((error) => {
      panel.webview.html = renderPreviewError(String(error), document);
    });
  }, delay ?? configuredDelay);

  previewTimers.set(key, timer);
}

async function updatePreviewPanel(context: vscode.ExtensionContext, document: vscode.TextDocument, panel: vscode.WebviewPanel, showLoading: boolean): Promise<void> {
  const key = document.uri.toString();
  const requestedVersion = document.version;
  previewVersions.set(key, requestedVersion);
  const outputDirectory = await resolveOutputDirectory(document);
  const outputLabel = previewOutputLabel(document, outputDirectory);

  if (showLoading) {
    panel.webview.html = renderPreviewLoading(document, outputLabel);
  }
  const result = await runCli(context, document, 'preview');
  if (previewVersions.get(key) !== requestedVersion) {
    return;
  }

  if (result.code !== 0) {
    panel.webview.html = renderPreviewError(result.stderr || 'OfficeIMO Markup preview failed.', document, outputLabel);
    return;
  }

  let envelope: MarkupEnvelope;
  try {
    envelope = parseJson<MarkupEnvelope>(result.stdout);
  } catch (error) {
    panel.webview.html = renderPreviewError(`${String(error)}\n\n${result.stderr || result.stdout}`, document, outputLabel);
    return;
  }

  const renderMermaidPreview = vscode.workspace
    .getConfiguration('officeimoMarkup')
    .get<boolean>('renderMermaidInPreview', true);
  const mermaidPreviewScript = panel.webview.asWebviewUri(vscode.Uri.joinPath(context.extensionUri, 'out', 'mermaidPreview.js')).toString();
  panel.webview.html = renderPreview(envelope.Document, envelope.Diagnostics ?? [], renderMermaidPreview, outputLabel, panel.webview, document, mermaidPreviewScript);
  updateDiagnostics(document, envelope.Diagnostics ?? []);
}

function previewLocalResourceRoots(context: vscode.ExtensionContext, document: vscode.TextDocument): vscode.Uri[] {
  const roots: vscode.Uri[] = [vscode.Uri.joinPath(context.extensionUri, 'out'), vscode.Uri.file(path.dirname(document.fileName))];
  for (const folder of vscode.workspace.workspaceFolders ?? []) {
    roots.push(folder.uri);
  }

  return roots;
}

async function validateDocument(context: vscode.ExtensionContext, document: vscode.TextDocument, showMessage: boolean): Promise<void> {
  const result = await runCli(context, document, 'validate');
  if (result.code !== 0 && !result.stdout.trim()) {
    diagnostics.set(document.uri, [new vscode.Diagnostic(new vscode.Range(0, 0, 0, 1), result.stderr || 'OfficeIMO Markup validation failed.', vscode.DiagnosticSeverity.Warning)]);
    if (showMessage) {
      vscode.window.showErrorMessage(result.stderr || 'OfficeIMO Markup validation failed.');
    }
    return;
  }

  if (!result.stdout.trim()) {
    diagnostics.set(document.uri, [new vscode.Diagnostic(new vscode.Range(0, 0, 0, 1), result.stderr || 'OfficeIMO Markup validation failed.', vscode.DiagnosticSeverity.Warning)]);
    return;
  }

  const envelope = parseJson<ValidationEnvelope>(result.stdout);
  updateDiagnostics(document, envelope.Diagnostics ?? []);
  if (showMessage) {
    const count = envelope.Diagnostics?.length ?? 0;
    vscode.window.showInformationMessage(count === 0 ? 'OfficeIMO Markup validation passed.' : `OfficeIMO Markup validation returned ${count} diagnostic(s).`);
  }
}

function updateDiagnostics(document: vscode.TextDocument, items: MarkupDiagnostic[]): void {
  diagnostics.set(document.uri, items.map((item) => {
    const range = diagnosticRange(document, item);
    const diagnostic = new vscode.Diagnostic(range, item.Message ?? 'OfficeIMO Markup diagnostic.', toDiagnosticSeverity(item.Severity));
    diagnostic.source = item.NodeKind ? `OfficeIMO Markup (${item.NodeKind})` : 'OfficeIMO Markup';
    return diagnostic;
  }));
}

function diagnosticRange(document: vscode.TextDocument, item: MarkupDiagnostic): vscode.Range {
  const text = document.getText();
  const sourceText = (item.NodeSourceText ?? '').trim();
  const sourceIndex = sourceText.length > 0 ? text.indexOf(sourceText) : -1;
  if (sourceIndex >= 0) {
    return lineRangeAt(document, sourceIndex);
  }

  const firstSourceLine = sourceText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line.length > 0);
  if (firstSourceLine) {
    const lineIndex = text.indexOf(firstSourceLine);
    if (lineIndex >= 0) {
      return lineRangeAt(document, lineIndex);
    }
  }

  const token = diagnosticNodeToken(item.NodeKind);
  if (token) {
    const tokenIndex = text.toLowerCase().indexOf(token);
    if (tokenIndex >= 0) {
      return lineRangeAt(document, tokenIndex);
    }
  }

  return new vscode.Range(0, 0, 0, Math.max(1, document.lineAt(0).text.length));
}

function lineRangeAt(document: vscode.TextDocument, offset: number): vscode.Range {
  const position = document.positionAt(offset);
  const line = document.lineAt(position.line);
  return new vscode.Range(position.line, 0, position.line, Math.max(1, line.text.length));
}

function diagnosticNodeToken(nodeKind?: string): string | undefined {
  switch ((nodeKind ?? '').toLowerCase()) {
    case 'slide':
      return '@slide';
    case 'section':
      return '@section';
    case 'sheet':
      return '@sheet';
    case 'range':
      return '::range';
    case 'formula':
      return '::formula';
    case 'namedtable':
      return '::table';
    case 'chart':
      return '::chart';
    case 'formatting':
      return '::format';
    case 'textbox':
      return '::textbox';
    case 'columns':
      return '::columns';
    case 'column':
      return '::column';
    case 'card':
      return '::card';
    case 'tableofcontents':
      return '::toc';
    case 'headerfooter':
      return '::header';
    case 'pagebreak':
      return '::pagebreak';
    default:
      return undefined;
  }
}

function toDiagnosticSeverity(value?: string): vscode.DiagnosticSeverity {
  switch ((value ?? '').toLowerCase()) {
    case 'error':
      return vscode.DiagnosticSeverity.Error;
    case 'warning':
      return vscode.DiagnosticSeverity.Warning;
    case 'info':
      return vscode.DiagnosticSeverity.Information;
    default:
      return vscode.DiagnosticSeverity.Hint;
  }
}

async function runCli(context: vscode.ExtensionContext, document: vscode.TextDocument, command: string, extraArgs: string[] = []): Promise<CliResult> {
  const profile = inferProfile(document);
  const invocation = resolveCliInvocation(context, command, ['--stdin', '--profile', profile, ...extraArgs]);
  const cwd = documentWorkingDirectory(context, document);

  const result = await runProcess(invocation.executable, invocation.args, cwd, document.getText());
  return result;
}

function runProcess(executable: string, args: string[], cwd: string, stdinText?: string): Promise<ProcessResult> {
  return new Promise((resolve) => {
    const child = spawn(executable, args, {
      cwd,
      windowsHide: true
    });

    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk: Buffer) => stdout += chunk.toString());
    child.stderr.on('data', (chunk: Buffer) => stderr += chunk.toString());
    child.on('close', (code) => resolve({ stdout, stderr, code }));
    child.on('error', (error) => resolve({ stdout, stderr: String(error), code: 1 }));
    if (stdinText === undefined) {
      child.stdin.end();
    } else {
      child.stdin.end(stdinText);
    }
  });
}

async function rememberMermaidRenderer(renderer: string): Promise<void> {
  await vscode.workspace
    .getConfiguration('officeimoMarkup')
    .update('mermaidCliPath', renderer, vscode.ConfigurationTarget.Global);
}

function findLocalMermaidRenderer(context: vscode.ExtensionContext): string | undefined {
  const candidates = localMermaidRendererCandidates(context);
  return candidates.find((candidate) => fs.existsSync(candidate));
}

function localMermaidRendererCandidates(context: vscode.ExtensionContext): string[] {
  const binDirectory = path.join(localMermaidInstallRoot(context), 'node_modules', '.bin');
  const candidates = process.platform === 'win32'
    ? ['mmdc.cmd', 'mmdc.exe', 'mmdc.ps1', 'mmdc']
    : ['mmdc'];

  return candidates.map((candidate) => path.join(binDirectory, candidate));
}

function localMermaidInstallRoot(context: vscode.ExtensionContext): string {
  return path.join(context.globalStorageUri.fsPath, 'tools', 'mermaid');
}

function resolveCliInvocation(context: vscode.ExtensionContext, command: string, args: string[]): { executable: string; args: string[] } {
  const configPath = vscode.workspace.getConfiguration('officeimoMarkup').get<string>('cliPath', '').trim();
  const cliPath = resolveConfiguredCliPath(configPath) ?? resolveDefaultCliPath(context);
  const commandArgs = [command, ...args];

  if (cliPath.endsWith('.csproj')) {
    const builtDll = findBuiltCliDll(cliPath);
    if (builtDll) {
      return { executable: 'dotnet', args: [builtDll, ...commandArgs] };
    }

    return { executable: 'dotnet', args: ['run', '--project', cliPath, '--', ...commandArgs] };
  }

  if (cliPath.endsWith('.dll')) {
    return { executable: 'dotnet', args: [cliPath, ...commandArgs] };
  }

  return { executable: cliPath, args: commandArgs };
}

function resolveConfiguredCliPath(configPath: string): string | undefined {
  if (!configPath) {
    return undefined;
  }

  const resolvedPath = resolveRealPath(path.resolve(configPath));
  if (!fs.existsSync(resolvedPath)) {
    return undefined;
  }

  if (isAllowedCliArtifact(resolvedPath)) {
    return resolvedPath;
  }

  if (!warnedAboutInvalidCliPath) {
    warnedAboutInvalidCliPath = true;
    void vscode.window.showWarningMessage(
      'officeimoMarkup.cliPath must point to an OfficeIMO.Markup.Cli executable, DLL, or csproj. Falling back to the bundled CLI.'
    );
  }

  return undefined;
}

function resolveDefaultCliPath(context: vscode.ExtensionContext): string {
  const extensionPath = resolveRealPath(context.extensionUri.fsPath);
  const bundledDll = path.join(extensionPath, 'tools', 'OfficeIMO.Markup.Cli', 'OfficeIMO.Markup.Cli.dll');
  const candidates = [
    bundledDll,
    path.join(workspaceRoot(context), 'OfficeIMO.Markup.Cli', 'OfficeIMO.Markup.Cli.csproj'),
    path.resolve(extensionPath, '..', 'OfficeIMO.Markup.Cli', 'OfficeIMO.Markup.Cli.csproj')
  ];

  return candidates.find((candidate) => fs.existsSync(candidate)) ?? bundledDll;
}

function findBuiltCliDll(csprojPath: string): string | undefined {
  const projectDirectory = path.dirname(csprojPath);
  const candidates = [
    path.join(projectDirectory, 'bin', 'Release', 'net8.0', 'OfficeIMO.Markup.Cli.dll'),
    path.join(projectDirectory, 'bin', 'Debug', 'net8.0', 'OfficeIMO.Markup.Cli.dll')
  ];

  return candidates.find((candidate) => fs.existsSync(candidate));
}

function isAllowedCliArtifact(cliPath: string): boolean {
  return allowedCliArtifactNames.has(path.basename(cliPath).toLowerCase());
}

function workspaceRoot(context: vscode.ExtensionContext): string {
  return vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? path.resolve(resolveRealPath(context.extensionUri.fsPath), '..');
}

function documentWorkingDirectory(context: vscode.ExtensionContext, document: vscode.TextDocument): string {
  if (document.uri.scheme === 'file' && document.uri.fsPath) {
    return path.dirname(resolveRealPath(document.uri.fsPath));
  }

  return workspaceRoot(context);
}

function resolveRealPath(value: string): string {
  try {
    return fs.realpathSync(value);
  } catch {
    return value;
  }
}

function inferProfile(document: vscode.TextDocument): string {
  const match = /^---[\s\S]*?^profile:\s*([A-Za-z-]+)/m.exec(document.getText());
  if (match) {
    return match[1].toLowerCase();
  }

  return vscode.workspace.getConfiguration('officeimoMarkup').get<string>('defaultProfile', 'presentation');
}

function isMarkupDocument(document: vscode.TextDocument): boolean {
  return document.languageId === 'officeimo-markup'
    || document.fileName.endsWith('.omd')
    || document.fileName.endsWith('.office.md');
}

function isMarkupCommandCandidate(document: vscode.TextDocument): boolean {
  return isMarkupDocument(document)
    || document.languageId === 'markdown'
    || document.fileName.endsWith('.md');
}

function isOfficeMarkupContent(document: vscode.TextDocument): boolean {
  if (isMarkupDocument(document)) {
    return true;
  }

  if (document.languageId !== 'markdown' && !document.fileName.endsWith('.md')) {
    return false;
  }

  const text = document.getText();
  if (/^---\s*[\r\n][\s\S]*?^profile\s*:\s*(presentation|document|workbook|common)\s*$/im.test(text)) {
    return true;
  }

  return /(^|\r?\n)\s*@(?:slide|section|sheet)\b/i.test(text)
    || /(^|\r?\n)\s*::(?:notes|textbox|image|chart|mermaid|columns|column|card|toc|header|footer|range|formula|table)\b/i.test(text);
}

async function resolveDocumentProfile(context: vscode.ExtensionContext, document: vscode.TextDocument): Promise<string | undefined> {
  const result = await runCli(context, document, 'preview');
  if (result.code === 0 && result.stdout.trim().length > 0) {
    try {
      return parseJson<MarkupEnvelope>(result.stdout).Document?.Profile?.toLowerCase();
    } catch {
      // Fall through to the cheap front-matter scan below.
    }
  }

  const match = /^---\s*[\r\n][\s\S]*?^profile\s*:\s*([A-Za-z-]+)\s*$/im.exec(document.getText());
  return match?.[1]?.toLowerCase();
}

function parseJson<T>(text: string): T {
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error('OfficeIMO Markup CLI returned invalid JSON.');
  }
}

function renderPreview(
  document: MarkupDocument | undefined,
  items: MarkupDiagnostic[],
  renderMermaidPreview: boolean,
  outputLabel: string,
  webview: vscode.Webview,
  sourceDocument: vscode.TextDocument,
  mermaidPreviewScript?: string
): string {
  const blocks = document?.Blocks ?? [];
  const slides = blocks.filter((block) => block.Kind === 'Slide');
  const profile = (document?.Profile ?? '').toLowerCase();
  const body = slides.length > 0
    ? slides.map((slide, index) => renderSlide(slide, index + 1, document, webview, sourceDocument)).join('')
    : profile === 'workbook'
      ? renderWorkbook(blocks)
      : profile === 'document'
        ? renderDocument(blocks)
    : blocks.map((block) => renderBlock(block)).join('');

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body { margin: 0; padding: 20px; font-family: var(--vscode-font-family); color: var(--vscode-foreground); background: var(--vscode-editor-background); }
    .toolbar { margin-bottom: 16px; display: flex; align-items: center; justify-content: space-between; gap: 12px; }
    .toolbar-title { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
    .toolbar-heading { color: var(--vscode-descriptionForeground); }
    .toolbar-subtitle { color: var(--vscode-descriptionForeground); font-size: 12px; opacity: 0.9; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .toolbar-actions { display: flex; flex-wrap: wrap; gap: 8px; }
    .toolbar-button { border: 1px solid var(--vscode-button-border, var(--vscode-panel-border)); background: var(--vscode-button-secondaryBackground, var(--vscode-button-background)); color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground)); border-radius: 6px; padding: 5px 10px; cursor: pointer; font: inherit; }
    .toolbar-button:hover { background: var(--vscode-button-secondaryHoverBackground, var(--vscode-button-hoverBackground)); }
    .deck { display: flex; flex-direction: column; align-items: center; gap: 22px; }
    .slide { position: relative; width: min(100%, 1120px); aspect-ratio: 16 / 9; border: 1px solid var(--vscode-panel-border); background: #f8fafc; color: #172033; padding: 22px; box-sizing: border-box; overflow: hidden; border-radius: 6px; background-position: center; background-repeat: no-repeat; background-size: cover; }
    .slide::before { content: ""; position: absolute; z-index: 0; top: -6%; right: 9%; width: 14%; height: 114%; background: #dbe6fb; transform: skewX(-5deg); }
    .slide::after { content: ""; position: absolute; z-index: 0; top: 0; right: 8%; width: 8px; height: 100%; background: #2563eb; }
    .slide.has-explicit-background::before, .slide.has-explicit-background::after { display: none; }
    .slide > * { position: relative; z-index: 1; }
    .slide-bg-overlay { position: absolute; inset: 0; z-index: 0; pointer-events: none; }
    .slide-label { position: absolute; z-index: 2; top: 12px; right: 18px; color: rgba(23, 32, 51, 0.34); font-size: 11px; font-weight: 700; pointer-events: none; }
    .slide h2 { margin: 0 0 12px; font-size: 22px; }
    .meta { font-size: 12px; color: var(--vscode-descriptionForeground); margin-bottom: 12px; }
    .block { margin: 8px 0; }
    .slide .block[style*="position:absolute"] { margin: 0; box-sizing: border-box; overflow: hidden; }
    pre { white-space: pre-wrap; background: var(--vscode-textCodeBlock-background); padding: 10px; border-radius: 6px; }
    ul { padding-left: 20px; }
    .columns { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 18px; margin-top: 12px; }
    .column { min-width: 0; }
    .textbox { font-weight: 600; margin: 10px 0; }
    .card { border: 1px solid var(--vscode-panel-border); border-radius: 6px; padding: 10px; }
    .document { max-width: 860px; margin: 0 auto; }
    .page { min-height: 720px; background: var(--vscode-editorWidget-background); border: 1px solid var(--vscode-panel-border); border-radius: 6px; padding: 44px 56px; box-sizing: border-box; }
    .page-break { display: flex; align-items: center; gap: 10px; color: var(--vscode-descriptionForeground); margin: 22px 0; }
    .page-break::before, .page-break::after { content: ""; height: 1px; background: var(--vscode-panel-border); flex: 1; }
    .toc { border-left: 3px solid var(--vscode-button-background); padding: 8px 0 8px 12px; margin: 12px 0; color: var(--vscode-descriptionForeground); }
    .header-footer, .section-marker, .formula, .named-table, .formatting { color: var(--vscode-descriptionForeground); font-size: 12px; border: 1px solid var(--vscode-panel-border); border-radius: 6px; padding: 8px 10px; margin: 8px 0; }
    .workbook { display: grid; gap: 16px; }
    .sheet { border: 1px solid var(--vscode-panel-border); border-radius: 6px; background: var(--vscode-editorWidget-background); overflow: hidden; }
    .sheet-title { padding: 10px 12px; border-bottom: 1px solid var(--vscode-panel-border); font-weight: 700; }
    .sheet-body { padding: 12px; }
    .sheet-grid-wrap { overflow: auto; border: 1px solid var(--vscode-panel-border); border-radius: 6px; background: var(--vscode-editor-background); margin-bottom: 12px; }
    .sheet-grid { border-collapse: separate; border-spacing: 0; min-width: 100%; width: max-content; }
    .sheet-grid th, .sheet-grid td { border-right: 1px solid var(--vscode-panel-border); border-bottom: 1px solid var(--vscode-panel-border); padding: 6px 8px; box-sizing: border-box; }
    .sheet-grid th { background: var(--vscode-editor-background); color: var(--vscode-descriptionForeground); font-size: 11px; font-weight: 600; text-align: center; position: sticky; top: 0; z-index: 1; }
    .sheet-grid .corner { min-width: 38px; width: 38px; position: sticky; left: 0; z-index: 3; }
    .sheet-grid .row-header { min-width: 38px; width: 38px; position: sticky; left: 0; z-index: 2; }
    .sheet-grid .cell { min-width: 88px; max-width: 220px; vertical-align: top; background: var(--vscode-editorWidget-background); color: var(--vscode-foreground); }
    .sheet-grid .cell.empty { color: transparent; }
    .sheet-grid .table-header-cell { font-weight: 700; background: color-mix(in srgb, var(--vscode-editorWidget-background) 82%, var(--vscode-button-background) 18%); }
    .sheet-grid .formula-cell { font-family: var(--vscode-editor-font-family, var(--vscode-font-family)); }
    .sheet-grid .cell-content { overflow-wrap: anywhere; }
    .sheet-grid .cell-meta { display: block; margin-top: 4px; font-size: 10px; color: var(--vscode-descriptionForeground); }
    .sheet-grid .format-summary { display: flex; flex-wrap: wrap; gap: 6px; margin: 0 0 10px; }
    .sheet-grid .format-chip { border: 1px solid var(--vscode-panel-border); border-radius: 999px; padding: 2px 8px; font-size: 11px; color: var(--vscode-descriptionForeground); background: var(--vscode-editor-background); }
    .data-table { border-collapse: collapse; width: 100%; margin: 10px 0; }
    .data-table th, .data-table td { border: 1px solid var(--vscode-panel-border); padding: 6px 8px; text-align: left; }
    .data-table th { background: var(--vscode-editor-background); font-weight: 700; }
    .range-caption { color: var(--vscode-descriptionForeground); font-size: 12px; margin-top: 10px; }
    .chart { border: 1px solid var(--vscode-panel-border); border-radius: 6px; padding: 10px; }
    .chart-title { font-weight: 700; margin-bottom: 8px; }
    .chart-placeholder { border: 1px dashed var(--vscode-panel-border); border-radius: 6px; padding: 10px; color: var(--vscode-descriptionForeground); background: var(--vscode-textCodeBlock-background); }
    .chart-meta { display: flex; flex-wrap: wrap; gap: 6px; margin: 0 0 10px; }
    .chart-chip { border: 1px solid var(--vscode-panel-border); border-radius: 6px; padding: 2px 6px; font-size: 11px; color: var(--vscode-descriptionForeground); background: var(--vscode-editor-background); }
    .chart-row { display: grid; grid-template-columns: 72px 1fr 44px; gap: 8px; align-items: center; margin: 6px 0; }
    .chart-category { margin: 8px 0 10px; }
    .chart-category-name { font-weight: 700; margin-bottom: 4px; }
    .chart-series-row { display: grid; grid-template-columns: minmax(58px, 86px) 1fr 44px; gap: 8px; align-items: center; margin: 4px 0; }
    .chart-series-name { color: var(--vscode-descriptionForeground); font-size: 11px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .chart-legend { display: flex; flex-wrap: wrap; gap: 10px; margin: 0 0 8px; color: var(--vscode-descriptionForeground); font-size: 11px; }
    .chart-legend-item { display: inline-flex; align-items: center; gap: 4px; }
    .chart-swatch { width: 9px; height: 9px; border-radius: 2px; display: inline-block; }
    .chart-bar-track { height: 10px; background: var(--vscode-editor-background); border: 1px solid var(--vscode-panel-border); }
    .chart-bar { height: 100%; background: var(--vscode-button-background); }
    .chart-value { text-align: right; color: var(--vscode-descriptionForeground); }
    .slide .chart { display: flex; flex-direction: column; gap: 4px; background: #ffffff; border-color: #d9e2ef; box-shadow: 0 6px 14px rgba(15, 23, 42, 0.06); }
    .slide .chart-title { margin-bottom: 2px; }
    .slide .chart-meta { display: none; }
    .slide .presentation-chart svg { width: 100%; flex: 1; min-height: 0; display: block; }
    .slide .chart-legend { margin-bottom: 2px; }
    .slide .chart-category { margin: 2px 0 4px; }
    .slide .chart-category-name { margin-bottom: 2px; font-size: 12px; }
    .slide .chart-series-row { margin: 2px 0; }
    .slide .chart-row { margin: 3px 0; }
    .image-placeholder, .diagram { border: 1px solid #d9e2ef; border-radius: 6px; padding: 10px; background: #ffffff; color: #172033; box-shadow: 0 6px 14px rgba(15, 23, 42, 0.06); }
    .diagram-title { font-size: 11px; color: var(--vscode-descriptionForeground); margin-bottom: 6px; text-transform: uppercase; }
    .diagram-status { display: none; margin-top: 8px; font-size: 12px; color: var(--vscode-descriptionForeground); }
    .diagram.pending .mermaid { opacity: 0; }
    .diagram.pending .diagram-source { display: none; }
    .diagram.rendered .diagram-source { display: none; }
    .diagram.render-failed .mermaid { display: none; }
    .diagram.render-failed .diagram-status { display: block; }
    .diagram .mermaid { height: calc(100% - 22px); min-height: 0; display: flex; align-items: center; justify-content: center; }
    .diagram .mermaid svg { width: 100% !important; max-width: 100%; max-height: 100%; height: auto !important; display: block; }
    .diagnostic { border-left: 3px solid var(--vscode-editorWarning-foreground); padding-left: 8px; margin-bottom: 8px; }
  </style>
</head>
<body>
  <div class="toolbar">
    <div class="toolbar-title">
      <div class="toolbar-heading">OfficeIMO Markup preview - ${escapeHtml(document?.Profile ?? 'unknown profile')}</div>
      <div class="toolbar-subtitle">${escapeHtml(outputLabel)}</div>
    </div>
    ${renderPreviewActions()}
  </div>
  ${items.length > 0 ? `<section>${items.map(renderDiagnostic).join('')}</section>` : ''}
  <main class="deck">${body || '<p>No previewable blocks yet.</p>'}</main>
  ${renderPreviewScript(renderMermaidPreview ? mermaidPreviewScript : undefined)}
</body>
</html>`;
}

function renderDocument(blocks: MarkupBlock[]): string {
  return `<section class="document"><div class="page">${blocks.map((block) => renderBlock(block, 'document')).join('')}</div></section>`;
}

function renderWorkbook(blocks: MarkupBlock[]): string {
  const sheets = buildWorkbookPreview(blocks);
  return `<section class="workbook">${sheets.map((sheet) => renderWorkbookSheet(sheet)).join('')}</section>`;
}

function buildWorkbookPreview(blocks: MarkupBlock[]): WorkbookPreviewSheet[] {
  const sheets = new Map<string, WorkbookPreviewSheet>();
  const order: string[] = [];
  let currentName = 'Workbook';

  const getSheet = (name?: string): WorkbookPreviewSheet => {
    const resolvedName = name && name.trim().length > 0 ? name.trim() : currentName;
    const existing = sheets.get(resolvedName);
    if (existing) {
      return existing;
    }

    const created: WorkbookPreviewSheet = {
      name: resolvedName,
      cells: new Map<string, WorkbookPreviewCell>(),
      tables: [],
      charts: [],
      extras: []
    };
    sheets.set(resolvedName, created);
    order.push(resolvedName);
    return created;
  };

  for (const block of blocks) {
    if (block.Kind === 'Sheet') {
      currentName = block.Name?.trim() || 'Sheet';
      getSheet(currentName);
      continue;
    }

    if (block.Kind === 'Range') {
      const resolved = resolveWorkbookReference(block.Sheet, block.Address, currentName);
      const parsed = parseWorkbookCellAddress(resolved.reference);
      const sheet = getSheet(resolved.sheetName);
      if (!parsed) {
        sheet.extras.push(block);
        continue;
      }

      const rows = block.Rows ?? [];
      for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        for (let columnIndex = 0; columnIndex < rows[rowIndex].length; columnIndex++) {
          const cell = getWorkbookPreviewCell(sheet, parsed.row + rowIndex, parsed.column + columnIndex);
          cell.text = rows[rowIndex][columnIndex] ?? '';
        }
      }
      continue;
    }

    if (block.Kind === 'Formula') {
      const resolved = resolveWorkbookReference(block.Sheet, block.Cell, currentName);
      const parsed = parseWorkbookCellAddress(resolved.reference);
      const sheet = getSheet(resolved.sheetName);
      if (!parsed) {
        sheet.extras.push(block);
        continue;
      }

      const cell = getWorkbookPreviewCell(sheet, parsed.row, parsed.column);
      cell.formula = block.Expression ?? '';
      cell.text = block.Expression ?? '';
      continue;
    }

    if (block.Kind === 'Formatting') {
      const explicitSheet = block.Attributes?.sheet;
      const resolved = resolveWorkbookReference(explicitSheet, block.Target, currentName);
      const cells = enumerateWorkbookCells(resolved.reference);
      const sheet = getSheet(resolved.sheetName);
      if (cells.length === 0) {
        sheet.extras.push(block);
        continue;
      }

      for (const cellAddress of cells) {
        const cell = getWorkbookPreviewCell(sheet, cellAddress.row, cellAddress.column);
        applyWorkbookFormatting(cell.style, block);
      }
      continue;
    }

    if (block.Kind === 'NamedTable') {
      const resolved = resolveWorkbookReference(block.Attributes?.sheet, block.Range, currentName);
      const range = parseWorkbookRangeReference(resolved.reference);
      const sheet = getSheet(resolved.sheetName);
      if (!range) {
        sheet.extras.push(block);
        continue;
      }

      sheet.tables.push({
        name: block.Name,
        hasHeader: block.HasHeader !== false,
        startRow: range.startRow,
        startColumn: range.startColumn,
        endRow: range.endRow,
        endColumn: range.endColumn
      });
      continue;
    }

    if (block.Kind === 'Chart') {
      const resolved = resolveWorkbookReference(block.Sheet, block.Attributes?.cell, currentName);
      getSheet(resolved.sheetName).charts.push(block);
      continue;
    }

    getSheet(currentName).extras.push(block);
  }

  if (order.length === 0) {
    getSheet(currentName);
  }

  return order.map((name) => sheets.get(name)!).filter(Boolean);
}

function renderWorkbookSheet(sheet: WorkbookPreviewSheet): string {
  const parts: string[] = [];
  const grid = renderWorkbookSheetGrid(sheet);
  if (grid) {
    parts.push(grid);
  }

  if (sheet.tables.length > 0) {
    parts.push(`<div class="format-summary">${sheet.tables.map((table) =>
      `<span class="format-chip">Table ${escapeHtml(table.name ?? '')}${table.name ? '' : ''} ${escapeHtml(`${columnName(table.startColumn)}${table.startRow}:${columnName(table.endColumn)}${table.endRow}`)}${table.hasHeader ? ' | header' : ''}</span>`
    ).join('')}</div>`);
  }

  parts.push(...sheet.charts.map((block) => renderBlock(block, 'workbook')));
  parts.push(...sheet.extras.map((block) => renderBlock(block, 'workbook')));

  const body = parts.length > 0
    ? parts.join('')
    : '<div class="range-caption">No previewable workbook content yet.</div>';

  return `<section class="sheet"><div class="sheet-title">${escapeHtml(sheet.name)}</div><div class="sheet-body">${body}</div></section>`;
}

function renderWorkbookSheetGrid(sheet: WorkbookPreviewSheet): string {
  if (sheet.cells.size === 0) {
    return '';
  }

  const addresses = Array.from(sheet.cells.keys()).map(parseWorkbookCellKey);
  const minRow = Math.min(...addresses.map((item) => item.row));
  const maxRow = Math.max(...addresses.map((item) => item.row));
  const minColumn = Math.min(...addresses.map((item) => item.column));
  const maxColumn = Math.max(...addresses.map((item) => item.column));
  const columnHeaders = Array.from({ length: maxColumn - minColumn + 1 }, (_, index) => minColumn + index);

  const headerRow = `<tr><th class="corner"></th>${columnHeaders.map((column) => `<th>${escapeHtml(columnName(column))}</th>`).join('')}</tr>`;
  const bodyRows = Array.from({ length: maxRow - minRow + 1 }, (_, rowOffset) => {
    const row = minRow + rowOffset;
    const cells = columnHeaders.map((column) => renderWorkbookPreviewCell(sheet, row, column)).join('');
    return `<tr><th class="row-header">${row}</th>${cells}</tr>`;
  }).join('');

  return `<div class="sheet-grid-wrap"><table class="sheet-grid"><thead>${headerRow}</thead><tbody>${bodyRows}</tbody></table></div>`;
}

function renderWorkbookPreviewCell(sheet: WorkbookPreviewSheet, row: number, column: number): string {
  const cell = sheet.cells.get(workbookCellKey(row, column));
  const classes = ['cell'];
  if (!cell || (!cell.text && !cell.formula)) {
    classes.push('empty');
  }

  if (cell?.formula) {
    classes.push('formula-cell');
  }

  if (isWorkbookHeaderCell(sheet, row, column)) {
    classes.push('table-header-cell');
  }

  const style = workbookPreviewCellStyle(cell?.style);
  const text = cell?.formula || cell?.text || '&nbsp;';
  const meta = cell?.style.numberFormat
    ? `<span class="cell-meta">${escapeHtml(cell.style.numberFormat)}</span>`
    : '';
  return `<td class="${classes.join(' ')}"${style}><div class="cell-content">${escapeHtml(text)}</div>${meta}</td>`;
}

function getWorkbookPreviewCell(sheet: WorkbookPreviewSheet, row: number, column: number): WorkbookPreviewCell {
  const key = workbookCellKey(row, column);
  let cell = sheet.cells.get(key);
  if (!cell) {
    cell = { style: {} };
    sheet.cells.set(key, cell);
  }

  return cell;
}

function applyWorkbookFormatting(style: WorkbookPreviewCellStyle, block: MarkupBlock): void {
  if (block.NumberFormat && block.NumberFormat.trim().length > 0) {
    style.numberFormat = block.NumberFormat;
  }

  const attributes = block.Attributes ?? {};
  const fill = getWorkbookAttribute(attributes, 'fill', 'background');
  const textColor = getWorkbookAttribute(attributes, 'color', 'font-color', 'fontColor', 'text-color', 'textColor', 'textcolor');
  const bold = getWorkbookAttribute(attributes, 'bold');
  const italic = getWorkbookAttribute(attributes, 'italic');
  const underline = getWorkbookAttribute(attributes, 'underline');
  const align = getWorkbookAttribute(attributes, 'align', 'alignment', 'horizontal-align', 'horizontalAlign', 'horizontalalignment', 'text-align', 'textAlign');
  const verticalAlign = getWorkbookAttribute(attributes, 'vertical-align', 'verticalAlign', 'verticalalignment', 'valign');
  const wrap = getWorkbookAttribute(attributes, 'wrap', 'wrap-text', 'wrapText');
  const border = getWorkbookAttribute(attributes, 'border', 'border-style', 'borderStyle');
  const borderColor = getWorkbookAttribute(attributes, 'border-color', 'borderColor', 'line-color', 'lineColor');

  const normalizedFill = cssColor(fill);
  if (normalizedFill) {
    style.fillColor = normalizedFill;
  }

  const normalizedTextColor = cssColor(textColor);
  if (normalizedTextColor) {
    style.textColor = normalizedTextColor;
  }

  if (bold !== undefined) {
    style.bold = isTruthyAttribute(bold);
  }

  if (italic !== undefined) {
    style.italic = isTruthyAttribute(italic);
  }

  if (underline !== undefined) {
    style.underline = isTruthyAttribute(underline);
  }

  const textAlign = cssTextAlign(align);
  if (textAlign) {
    style.textAlign = textAlign;
  }

  const normalizedVerticalAlign = cssVerticalAlign(verticalAlign);
  if (normalizedVerticalAlign) {
    style.verticalAlign = normalizedVerticalAlign;
  }

  const normalizedBorderStyle = cssBorderStyle(border);
  if (normalizedBorderStyle) {
    style.borderStyle = normalizedBorderStyle;
  }

  const normalizedBorderColor = cssColor(borderColor) ?? (style.borderStyle ? '#94a3b8' : undefined);
  if (normalizedBorderColor) {
    style.borderColor = normalizedBorderColor;
  }

  if (wrap !== undefined) {
    style.wrap = isTruthyAttribute(wrap);
  }
}

function workbookPreviewCellStyle(style?: WorkbookPreviewCellStyle): string {
  if (!style) {
    return '';
  }

  const parts: string[] = [];
  if (style.fillColor) {
    parts.push(`background:${style.fillColor}`);
  }

  if (style.textColor) {
    parts.push(`color:${style.textColor}`);
  }

  if (style.bold !== undefined) {
    parts.push(`font-weight:${style.bold ? 700 : 400}`);
  }

  if (style.italic !== undefined) {
    parts.push(`font-style:${style.italic ? 'italic' : 'normal'}`);
  }

  if (style.underline !== undefined) {
    parts.push(`text-decoration:${style.underline ? 'underline' : 'none'}`);
  }

  if (style.textAlign) {
    parts.push(`text-align:${style.textAlign}`);
  }

  if (style.verticalAlign) {
    parts.push(`vertical-align:${style.verticalAlign}`);
  }

  if (style.wrap !== undefined) {
    parts.push(`white-space:${style.wrap ? 'pre-wrap' : 'nowrap'}`);
  }

  if (style.borderStyle) {
    parts.push(`border:1px ${style.borderStyle} ${style.borderColor ?? '#94a3b8'}`);
  }

  return parts.length === 0 ? '' : ` style="${parts.join(';')}"`;
}

function resolveWorkbookReference(explicitSheet: string | undefined, reference: string | undefined, fallbackSheet: string): { sheetName: string; reference: string } {
  const split = splitWorkbookSheetQualifiedReference(reference);
  if (split) {
    return split;
  }

  const sheetName = explicitSheet && explicitSheet.trim().length > 0
    ? explicitSheet.trim()
    : fallbackSheet;
  return { sheetName, reference: reference?.trim() ?? '' };
}

function splitWorkbookSheetQualifiedReference(reference: string | undefined): { sheetName: string; reference: string } | undefined {
  if (!reference || reference.trim().length === 0) {
    return undefined;
  }

  const value = reference.trim();
  const bangIndex = value.lastIndexOf('!');
  if (bangIndex <= 0 || bangIndex >= value.length - 1) {
    return undefined;
  }

  const sheetName = value.substring(0, bangIndex).trim().replace(/^'(.*)'$/, '$1').replace(/''/g, '\'');
  const localReference = value.substring(bangIndex + 1).trim();
  if (!sheetName || !localReference) {
    return undefined;
  }

  return { sheetName, reference: localReference };
}

function parseWorkbookCellAddress(address: string | undefined): { row: number; column: number } | undefined {
  const value = address?.trim();
  if (!value) {
    return undefined;
  }

  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(value);
  if (!match) {
    return undefined;
  }

  return {
    column: columnNumber(match[1]),
    row: Number.parseInt(match[2], 10)
  };
}

function parseWorkbookRangeReference(reference: string | undefined): { startRow: number; startColumn: number; endRow: number; endColumn: number } | undefined {
  const value = reference?.trim();
  if (!value) {
    return undefined;
  }

  const parts = value.split(':').map((part) => part.trim());
  if (parts.length === 1) {
    const cell = parseWorkbookCellAddress(parts[0]);
    return cell
      ? { startRow: cell.row, startColumn: cell.column, endRow: cell.row, endColumn: cell.column }
      : undefined;
  }

  if (parts.length !== 2) {
    return undefined;
  }

  const start = parseWorkbookCellAddress(parts[0]);
  const end = parseWorkbookCellAddress(parts[1]);
  if (!start || !end) {
    return undefined;
  }

  return {
    startRow: Math.min(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endRow: Math.max(start.row, end.row),
    endColumn: Math.max(start.column, end.column)
  };
}

function enumerateWorkbookCells(reference: string | undefined): Array<{ row: number; column: number }> {
  const range = parseWorkbookRangeReference(reference);
  if (!range) {
    return [];
  }

  const cells: Array<{ row: number; column: number }> = [];
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      cells.push({ row, column });
    }
  }

  return cells;
}

function workbookCellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function parseWorkbookCellKey(key: string): { row: number; column: number } {
  const [rowText, columnText] = key.split(':');
  return {
    row: Number.parseInt(rowText, 10),
    column: Number.parseInt(columnText, 10)
  };
}

function columnNumber(name: string): number {
  let result = 0;
  for (const character of name.toUpperCase()) {
    result = (result * 26) + (character.charCodeAt(0) - 64);
  }

  return result;
}

function columnName(column: number): string {
  let current = column;
  let result = '';
  while (current > 0) {
    current--;
    result = String.fromCharCode(65 + (current % 26)) + result;
    current = Math.floor(current / 26);
  }

  return result || 'A';
}

function isWorkbookHeaderCell(sheet: WorkbookPreviewSheet, row: number, column: number): boolean {
  return sheet.tables.some((table) =>
    table.hasHeader &&
    row === table.startRow &&
    column >= table.startColumn &&
    column <= table.endColumn);
}

function getWorkbookAttribute(attributes: Record<string, string>, ...names: string[]): string | undefined {
  for (const name of names) {
    const value = attributes[name];
    if (value !== undefined) {
      return value;
    }
  }

  return undefined;
}

function isTruthyAttribute(value: string): boolean {
  const normalized = value.trim().toLowerCase();
  return normalized === 'true' || normalized === 'yes' || normalized === 'on' || normalized === '1';
}

function renderPreviewLoading(document: vscode.TextDocument, outputLabel: string): string {
  return renderShell(
    `OfficeIMO Markup preview - refreshing ${escapeHtml(path.basename(document.fileName))}`,
    '<div class="placeholder">Refreshing preview...</div>',
    outputLabel
  );
}

function renderPreviewError(message: string, document: vscode.TextDocument, outputLabel: string): string {
  return renderShell(
    `OfficeIMO Markup preview - ${escapeHtml(path.basename(document.fileName))}`,
    `<div class="error"><strong>Preview failed</strong><pre>${escapeHtml(message)}</pre></div>`,
    outputLabel
  );
}

function renderShell(title: string, body: string, outputLabel: string): string {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body { margin: 0; padding: 20px; font-family: var(--vscode-font-family); color: var(--vscode-foreground); background: var(--vscode-editor-background); }
    .toolbar { margin-bottom: 16px; display: flex; align-items: center; justify-content: space-between; gap: 12px; }
    .toolbar-title { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
    .toolbar-heading { color: var(--vscode-descriptionForeground); }
    .toolbar-subtitle { color: var(--vscode-descriptionForeground); font-size: 12px; opacity: 0.9; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .toolbar-actions { display: flex; flex-wrap: wrap; gap: 8px; }
    .toolbar-button { border: 1px solid var(--vscode-button-border, var(--vscode-panel-border)); background: var(--vscode-button-secondaryBackground, var(--vscode-button-background)); color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground)); border-radius: 6px; padding: 5px 10px; cursor: pointer; font: inherit; }
    .toolbar-button:hover { background: var(--vscode-button-secondaryHoverBackground, var(--vscode-button-hoverBackground)); }
    .placeholder { border: 1px dashed var(--vscode-panel-border); border-radius: 6px; padding: 18px; color: var(--vscode-descriptionForeground); }
    .error { border-left: 3px solid var(--vscode-editorError-foreground); padding-left: 12px; }
    pre { white-space: pre-wrap; background: var(--vscode-textCodeBlock-background); padding: 10px; border-radius: 6px; }
  </style>
</head>
<body>
  <div class="toolbar">
    <div class="toolbar-title">
      <div class="toolbar-heading">${title}</div>
      <div class="toolbar-subtitle">${escapeHtml(outputLabel)}</div>
    </div>
    ${renderPreviewActions()}
  </div>
  ${body}
  ${renderPreviewScript()}
</body>
</html>`;
}

function renderSlide(slide: MarkupBlock, index: number, markupDocument: MarkupDocument | undefined, webview: vscode.Webview, sourceDocument: vscode.TextDocument): string {
  const metadata = [
    slide.Section ? `section: ${slide.Section}` : '',
    slide.Layout ? `layout: ${slide.Layout}` : '',
    summarizePreviewTransition(slide),
    slide.Background ? `background: ${slide.Background}` : ''
  ].filter(Boolean).join(' | ');
  const blocks = slide.Blocks ?? [];
  const background = slidePreviewBackground(slide, markupDocument, webview, sourceDocument);

  const title = shouldRenderPreviewSlideTitle(slide)
    ? `<h2>${escapeHtml(slide.Title ?? `Slide ${index}`)}</h2>`
    : '';

  return `<section class="slide ${background.className}"${background.style}>
    ${background.overlayHtml}
    <div class="slide-label" title="${escapeHtml(metadata)}">Slide ${index}${slide.Section ? ` · ${escapeHtml(slide.Section)}` : ''}</div>
    ${title}
    ${renderSlideBlocks(blocks)}
  </section>`;
}

function shouldRenderPreviewSlideTitle(slide: MarkupBlock): boolean {
  if (!slide.Title || slide.Title.trim().length === 0) {
    return false;
  }

  return normalizeLayout(slide.Layout ?? '') !== 'blank';
}

function normalizeLayout(value: string): string {
  return value.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
}

function slidePreviewBackground(slide: MarkupBlock, markupDocument: MarkupDocument | undefined, webview: vscode.Webview, sourceDocument: vscode.TextDocument): SlidePreviewBackground {
  const background = (slide.Background ?? '').trim();
  if (background.length === 0) {
    return { className: '', style: '', overlayHtml: '' };
  }

  const theme = resolvePreviewTheme(markupDocument);
  const styles: string[] = [];
  const classNames = ['has-explicit-background'];
  const solid = extractFunctionArgument(background, 'solid');
  const gradient = extractFunctionArgument(background, 'gradient');
  const image = extractFunctionArgument(background, 'image');
  const fit = extractNamedValue(background, 'fit');
  const angle = extractNamedValue(background, 'angle');
  const overlay = extractNamedValue(background, 'overlay');

  if (solid) {
    const color = resolvePreviewBackgroundColor(solid, theme);
    if (color) {
      styles.push(`background:${color}`);
    }
  } else if (gradient) {
    const parts = gradient.split(',')
      .map((part) => resolvePreviewBackgroundColor(part.trim(), theme))
      .filter((part): part is string => Boolean(part));
    if (parts.length >= 2) {
      styles.push(`background:linear-gradient(${previewGradientAngle(angle)}deg, ${parts[0]}, ${parts[1]})`);
    }
  } else {
    const color = resolvePreviewBackgroundColor(background, theme);
    if (color) {
      styles.push(`background:${color}`);
    }
  }

  if (image) {
    const resolved = resolvePreviewAssetPath(image, sourceDocument);
    if (resolved && fs.existsSync(resolved)) {
      const uri = webview.asWebviewUri(vscode.Uri.file(resolved)).toString();
      styles.push(`background-image:url('${escapeHtmlAttribute(uri)}')`);
      styles.push(`background-size:${previewBackgroundSize(fit)}`);
    }
  }

  const overlayColor = resolvePreviewBackgroundColor(overlay, theme) ?? overlay;
  const overlayHtml = overlayColor
    ? `<div class="slide-bg-overlay" style="background:${escapeHtmlAttribute(overlayColor)}"></div>`
    : '';

  return {
    className: classNames.join(' '),
    style: styles.length > 0 ? ` style="${styles.map((value) => escapeHtmlAttribute(value)).join(';')}"` : '',
    overlayHtml
  };
}

function summarizePreviewTransition(slide: MarkupBlock): string | undefined {
  const details = slide.TransitionDetails;
  if (!details) {
    return slide.Transition ? `transition: ${slide.Transition}` : undefined;
  }

  const parts: string[] = [];
  if (details.ResolvedIdentifier) {
    parts.push(details.ResolvedIdentifier);
  } else if (details.Effect) {
    parts.push(details.Effect);
  } else if (details.RawText) {
    parts.push(details.RawText);
  }

  const attributes = details.Attributes ?? {};
  const direction = attributes.direction ?? attributes.dir ?? attributes.orientation ?? attributes.axis ?? attributes.mode;
  const duration = attributes.duration;
  const speed = attributes.speed ?? attributes.spd;
  const advanceOnClick = attributes['advance-on-click'] ?? attributes.advanceonclick ?? attributes['advance-click'] ?? attributes.onclick ?? attributes.click;
  const advanceAfter = attributes['advance-after'] ?? attributes.advanceafter ?? attributes.after ?? attributes.delay;
  if (direction) {
    parts.push(`direction=${direction}`);
  }

  if (duration) {
    parts.push(`duration=${duration}`);
  }

  if (speed) {
    parts.push(`speed=${speed}`);
  }

  if (advanceOnClick) {
    parts.push(`advance-on-click=${advanceOnClick}`);
  }

  if (advanceAfter) {
    parts.push(`advance-after=${advanceAfter}`);
  }

  return parts.length > 0
    ? `transition: ${parts.join(' ')}`
    : slide.Transition ? `transition: ${slide.Transition}` : undefined;
}

function resolvePreviewAssetPath(source: string, document: vscode.TextDocument): string | undefined {
  const cleaned = source.trim().replace(/^['"]|['"]$/g, '');
  if (!cleaned) {
    return undefined;
  }

  return path.isAbsolute(cleaned)
    ? cleaned
    : path.resolve(path.dirname(document.fileName), cleaned);
}

function previewBackgroundSize(value: string | undefined): string {
  const normalized = (value ?? '').trim().toLowerCase();
  switch (normalized) {
    case 'contain':
      return 'contain';
    case 'stretch':
      return '100% 100%';
    default:
      return 'cover';
  }
}

function previewGradientAngle(value: string | undefined): number {
  const trimmed = value?.trim();
  if (!trimmed) {
    return 135;
  }

  const normalized = trimmed.replace(/deg$/i, '').trim();
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : 135;
}

function resolvePreviewTheme(document: MarkupDocument | undefined): PreviewTheme {
  const metadata = document?.Metadata ?? {};
  const themeName = metadata.theme ?? '';
  const accent = normalizePreviewHexColor(
    metadata.accent
    ?? metadata['accent-color']
    ?? metadata['brand-color'])
    ?? defaultPreviewAccent(themeName);

  return {
    background: '#FFFFFF',
    surface: '#F7FAFC',
    panel: '#FFFFFF',
    panelBorder: '#D8E1E8',
    text: '#30343B',
    textSecondary: '#65717D',
    textMuted: '#8A949E',
    accent,
    accentDark: mixPreviewColor(accent, '#000000', 0.28),
    accentLight: mixPreviewColor(accent, '#FFFFFF', 0.82),
    accent2: mixPreviewColor(accent, '#FFFFFF', 0.18),
    accent3: defaultPreviewAccent3(themeName),
    warning: '#F4A100'
  };
}

function resolvePreviewBackgroundColor(value: string | undefined, theme: PreviewTheme): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) {
    return undefined;
  }

  if (/^rgba?\(/i.test(trimmed)) {
    return trimmed;
  }

  const color = cssColor(trimmed);
  if (color) {
    return color;
  }

  switch (trimmed.replace(/[^a-zA-Z0-9]/g, '').toLowerCase()) {
    case 'primary':
      return theme.accentDark;
    case 'accent':
    case 'accent1':
    case 'brand':
      return theme.accent;
    case 'accentdark':
      return theme.accentDark;
    case 'accentlight':
      return theme.accentLight;
    case 'accent2':
    case 'secondary':
      return theme.accent2;
    case 'accent3':
    case 'tertiary':
      return theme.accent3;
    case 'warning':
      return theme.warning;
    case 'background':
    case 'background1':
    case 'bg1':
      return theme.background;
    case 'surface':
    case 'background2':
    case 'bg2':
      return theme.surface;
    case 'panel':
      return theme.panel;
    case 'panelborder':
    case 'border':
      return theme.panelBorder;
    case 'text':
    case 'text1':
    case 'foreground':
      return theme.text;
    case 'text2':
    case 'secondarytext':
      return theme.textSecondary;
    case 'muted':
    case 'mutedtext':
      return theme.textMuted;
    case 'white':
      return '#FFFFFF';
    case 'black':
      return '#000000';
    default:
      return undefined;
  }
}

function defaultPreviewAccent(themeName: string): string {
  switch (themeName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase()) {
    case 'modernblue':
      return '#0098C8';
    case 'evotecmodern':
    default:
      return '#2563EB';
  }
}

function defaultPreviewAccent3(themeName: string): string {
  switch (themeName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase()) {
    case 'modernblue':
      return '#6F6AA6';
    case 'evotecmodern':
    default:
      return '#6F6AA6';
  }
}

function normalizePreviewHexColor(value: string | undefined): string | undefined {
  const color = cssColor(value);
  return color ? color.toUpperCase() : undefined;
}

function mixPreviewColor(source: string, target: string, ratio: number): string {
  const from = parsePreviewHexColor(source);
  const to = parsePreviewHexColor(target);
  const clamped = Math.max(0, Math.min(1, ratio));
  const red = Math.round(from.red + ((to.red - from.red) * clamped));
  const green = Math.round(from.green + ((to.green - from.green) * clamped));
  const blue = Math.round(from.blue + ((to.blue - from.blue) * clamped));
  return `#${red.toString(16).padStart(2, '0')}${green.toString(16).padStart(2, '0')}${blue.toString(16).padStart(2, '0')}`.toUpperCase();
}

function parsePreviewHexColor(value: string): { red: number; green: number; blue: number } {
  const normalized = value.startsWith('#') ? value.substring(1) : value;
  return {
    red: Number.parseInt(normalized.substring(0, 2), 16),
    green: Number.parseInt(normalized.substring(2, 4), 16),
    blue: Number.parseInt(normalized.substring(4, 6), 16)
  };
}

function extractFunctionArgument(value: string, functionName: string): string | undefined {
  const expression = new RegExp(`${functionName}\\(([^)]*)\\)`, 'i');
  const match = value.match(expression);
  return match?.[1]?.trim().replace(/^['"]|['"]$/g, '') || undefined;
}

function extractNamedValue(value: string, attributeName: string): string | undefined {
  const index = value.toLowerCase().indexOf(`${attributeName.toLowerCase()}=`);
  if (index < 0) {
    return undefined;
  }

  const start = index + attributeName.length + 1;
  const remaining = value.substring(start);
  if (/^rgba\(/i.test(remaining)) {
    const end = remaining.indexOf(')');
    return end >= 0 ? remaining.substring(0, end + 1).trim() : undefined;
  }

  const nextSpace = remaining.indexOf(' ');
  return (nextSpace >= 0 ? remaining.substring(0, nextSpace) : remaining).trim() || undefined;
}

function renderSlideBlocks(blocks: MarkupBlock[]): string {
  let html = '';
  for (let index = 0; index < blocks.length; index++) {
    const block = blocks[index];
    if (isColumns(block)) {
      const columns: string[] = [];
      index++;
      while (index < blocks.length) {
        const current = blocks[index];
        if (!isColumn(current)) {
          index--;
          break;
        }

        const columnBlocks: MarkupBlock[] = [];
        const body = columnBody(current);
        if (body) {
          columnBlocks.push(...blocksFromMarkdownishBody(body));
        }

        index++;
        while (index < blocks.length && !isColumn(blocks[index])) {
          columnBlocks.push(blocks[index]);
          index++;
        }

        index--;
        columns.push(`<div class="column">${columnBlocks.map((columnBlock) => renderBlock(columnBlock, 'slide')).join('')}</div>`);
        index++;
      }

      index--;
      html += `<div class="columns"${blockStyle(block)}>${columns.join('')}</div>`;
      continue;
    }

    html += renderBlock(block, 'slide');
  }

  return html;
}

function renderBlock(block: MarkupBlock, mode: PreviewRenderMode = 'document'): string {
  switch (block.Kind) {
    case 'Paragraph':
      return `<p class="block"${blockStyle(block)}>${escapeHtml(block.Text ?? '')}</p>`;
    case 'Heading':
      return `<h3 class="block"${blockStyle(block)}>${escapeHtml(block.Text ?? '')}</h3>`;
    case 'List':
      return `<ul class="block"${blockStyle(block)}>${(block.Items ?? []).map((item) => `<li>${escapeHtml(item.Text ?? '')}</li>`).join('')}</ul>`;
    case 'Code':
      return `<pre class="block"${blockStyle(block)}>${escapeHtml(block.Content ?? '')}</pre>`;
    case 'Image':
      return `<div class="block image-placeholder"${blockStyle(block)}>Image: ${escapeHtml(block.Source ?? '')}</div>`;
    case 'Table':
      return renderDataTable(block.Headers ?? [], block.Rows ?? []);
    case 'Diagram':
      return renderDiagram(block);
    case 'Chart':
      return renderChart(block, mode);
    case 'TextBox':
      return `<div class="block textbox"${blockStyle(block)}>${renderInlineMarkdownish(block.Text ?? '')}</div>`;
    case 'Card':
      return renderCard(block);
    case 'Columns':
    case 'Column':
      return '';
    case 'TableOfContents':
      return `<div class="block toc">Table of contents${block.MinLevel || block.MaxLevel ? `, levels ${escapeHtml(String(block.MinLevel ?? 1))}-${escapeHtml(String(block.MaxLevel ?? 6))}` : ''}</div>`;
    case 'HeaderFooter':
      return `<div class="block header-footer">${escapeHtml(block.Name ?? 'header/footer')}: ${escapeHtml(block.Text ?? '')}</div>`;
    case 'PageBreak':
      return '<div class="block page-break">page break</div>';
    case 'Section':
      return `<div class="block section-marker">Section${block.Name ? `: ${escapeHtml(block.Name)}` : ''}${block.Orientation ? ` | orientation: ${escapeHtml(block.Orientation)}` : ''}</div>`;
    case 'Sheet':
      return `<div class="block section-marker">Sheet: ${escapeHtml(block.Name ?? 'Sheet')}</div>`;
    case 'Range':
      return `<div class="block"><div class="range-caption">${escapeHtml(block.Sheet ? `${block.Sheet}!` : '')}${escapeHtml(block.Address ?? 'range')}</div>${renderDataTable([], block.Rows ?? [])}</div>`;
    case 'NamedTable':
      return `<div class="block named-table">Named table ${escapeHtml(block.Name ?? '')} | ${escapeHtml(block.Range ?? '')}${block.HasHeader === false ? ' | no header' : ''}</div>`;
    case 'Formula':
      return `<div class="block formula">${escapeHtml(block.Cell ?? '')}: ${escapeHtml(block.Expression ?? '')}</div>`;
    case 'Formatting':
      return `<div class="block formatting">Format ${escapeHtml(block.Target ?? '')}${block.NumberFormat ? ` | ${escapeHtml(block.NumberFormat)}` : ''}${block.Style ? ` | ${escapeHtml(block.Style)}` : ''}</div>`;
    case 'RawMarkdown':
      return `<pre class="block">${escapeHtml(block.Markdown ?? '')}</pre>`;
    case 'Extension':
      return renderExtension(block, mode);
    default:
      return `<div class="block">${escapeHtml(block.Kind ?? 'Block')}</div>`;
  }
}

function renderDataTable(headers: string[], rows: string[][]): string {
  const hasHeaders = headers.length > 0;
  const bodyRows = hasHeaders ? rows : rows.slice(1);
  const inferredHeaders = hasHeaders ? headers : (rows[0] ?? []);
  const head = inferredHeaders.length > 0
    ? `<thead><tr>${inferredHeaders.map((cell) => `<th>${escapeHtml(cell)}</th>`).join('')}</tr></thead>`
    : '';
  const body = bodyRows.length > 0
    ? `<tbody>${bodyRows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`).join('')}</tbody>`
    : '';

  return `<table class="data-table">${head}${body}</table>`;
}

function renderDiagram(block: MarkupBlock): string {
  const language = (block.Language ?? 'diagram').toLowerCase();
  const content = block.Content ?? '';
  if (language === 'mermaid') {
    return `<div class="block diagram mermaid-diagram pending" data-language="mermaid"${blockStyle(block)}>
      <div class="diagram-title">Mermaid</div>
      <div class="mermaid">${escapeHtml(content)}</div>
      <pre class="diagram-source">${escapeHtml(content)}</pre>
      <div class="diagram-status">Mermaid preview unavailable. Raw source is shown.</div>
    </div>`;
  }

  return `<div class="block diagram"${blockStyle(block)}><div class="diagram-title">${escapeHtml(block.Language ?? 'diagram')}</div><pre>${escapeHtml(content)}</pre></div>`;
}

function renderExtension(block: MarkupBlock, mode: PreviewRenderMode = 'document'): string {
  const command = (block.Command ?? 'extension').toLowerCase();
  switch (command) {
    case 'textbox':
      return `<div class="block textbox"${blockStyle(block)}>${renderInlineMarkdownish(block.Body ?? '')}</div>`;
    case 'column':
    case 'left':
    case 'right':
      return `<div class="block column">${renderMarkdownish(block.Body ?? '', mode)}</div>`;
    case 'columns':
      return '';
    default:
      return `<div class="block">${escapeHtml(block.Command ?? 'extension')}: ${renderInlineMarkdownish(block.Body ?? '')}</div>`;
  }
}

function renderCard(block: MarkupBlock): string {
  const title = block.Title ? `<strong>${escapeHtml(block.Title)}</strong><br>` : '';
  return `<div class="block card"${blockStyle(block)}>${title}${renderInlineMarkdownish(block.Body ?? '')}</div>`;
}

function renderChart(block: MarkupBlock, mode: PreviewRenderMode = 'document'): string {
  const rows = block.Rows ?? [];
  const title = block.Title ?? block.ChartType ?? 'chart';
  const meta = mode === 'slide' ? '' : renderChartMeta(block);
  if (rows.length < 2 || rows[0].length < 2) {
    const body = block.Source
      ? renderSourceBackedChartPlaceholder(block)
      : '<div class="meta">No inline chart data yet.</div>';
    return `<div class="block chart"${blockStyle(block)}><div class="chart-title">${escapeHtml(title)}</div>${meta}${body}</div>`;
  }

  const seriesNames = rows[0].slice(1).map((name, index) => name || `Series ${index + 1}`);
  const dataRows = rows.slice(1).map((row) => ({
    category: row[0] ?? '',
    values: seriesNames.map((_, index) => Number.parseFloat(row[index + 1] ?? '0'))
  })).filter((row) => row.category.length > 0);
  const numericValues = dataRows.flatMap((row) => row.values).filter((value) => Number.isFinite(value) && value > 0);
  const max = Math.max(...numericValues, 1);

  if (mode === 'slide') {
    return renderSlideChart(block, title, seriesNames, dataRows, max);
  }

  if (seriesNames.length <= 1) {
    const bars = dataRows.map((row) => {
      const value = Number.isFinite(row.values[0]) ? row.values[0] : 0;
      const width = chartBarWidth(value, max);
      return `<div class="chart-row">
        <div>${escapeHtml(row.category)}</div>
        <div class="chart-bar-track"><div class="chart-bar" style="width:${width.toFixed(1)}%"></div></div>
        <div class="chart-value">${escapeHtml(formatChartValue(value))}</div>
      </div>`;
    }).join('');

    return `<div class="block chart"${blockStyle(block)}><div class="chart-title">${escapeHtml(title)}</div>${meta}${bars}</div>`;
  }

  const legend = `<div class="chart-legend">${seriesNames.map((name, index) =>
    `<span class="chart-legend-item"><span class="chart-swatch" style="background:${chartColor(index)}"></span>${escapeHtml(name)}</span>`
  ).join('')}</div>`;
  const bars = dataRows.map((row) => {
    const series = row.values.map((value, index) => {
      const safeValue = Number.isFinite(value) ? value : 0;
      const width = chartBarWidth(safeValue, max);
      return `<div class="chart-series-row">
        <div class="chart-series-name">${escapeHtml(seriesNames[index])}</div>
        <div class="chart-bar-track"><div class="chart-bar" style="width:${width.toFixed(1)}%;background:${chartColor(index)}"></div></div>
        <div class="chart-value">${escapeHtml(formatChartValue(safeValue))}</div>
      </div>`;
    }).join('');

    return `<div class="chart-category">
      <div class="chart-category-name">${escapeHtml(row.category)}</div>
      ${series}
    </div>`;
  }).join('');

  return `<div class="block chart"${blockStyle(block)}><div class="chart-title">${escapeHtml(title)}</div>${meta}${legend}${bars}</div>`;
}

function renderSlideChart(block: MarkupBlock, title: string, seriesNames: string[], dataRows: ChartPreviewRow[], max: number): string {
  const svgWidth = 800;
  const svgHeight = 330;
  const left = 58;
  const right = 24;
  const top = 22;
  const bottom = 42;
  const plotWidth = svgWidth - left - right;
  const plotHeight = svgHeight - top - bottom;
  const groupWidth = plotWidth / Math.max(1, dataRows.length);
  const seriesCount = Math.max(1, seriesNames.length);
  const barWidth = Math.max(10, Math.min(28, (groupWidth * 0.62) / seriesCount));
  const groupBarWidth = barWidth * seriesCount;
  const gridLines = [0.25, 0.5, 0.75, 1];
  const grid = gridLines.map((line) => {
    const y = top + plotHeight - (plotHeight * line);
    return `<line x1="${left}" y1="${y.toFixed(1)}" x2="${svgWidth - right}" y2="${y.toFixed(1)}" stroke="#e5e7eb" stroke-width="1" />`;
  }).join('');
  const bars = dataRows.map((row, rowIndex) => {
    const groupLeft = left + (rowIndex * groupWidth) + ((groupWidth - groupBarWidth) / 2);
    const category = `<text x="${(left + rowIndex * groupWidth + groupWidth / 2).toFixed(1)}" y="${svgHeight - 14}" text-anchor="middle" font-size="12" fill="#475569">${escapeHtml(row.category)}</text>`;
    const rects = row.values.map((value, seriesIndex) => {
      const safeValue = Number.isFinite(value) ? Math.max(0, value) : 0;
      const height = (safeValue / max) * plotHeight;
      const x = groupLeft + (seriesIndex * barWidth);
      const y = top + plotHeight - height;
      return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${Math.max(4, barWidth - 3).toFixed(1)}" height="${height.toFixed(1)}" rx="2" fill="${chartColor(seriesIndex)}"><title>${escapeHtml(row.category)} ${escapeHtml(seriesNames[seriesIndex])}: ${escapeHtml(formatChartValue(safeValue))}</title></rect>`;
    }).join('');

    return rects + category;
  }).join('');
  const legend = seriesNames.map((name, index) => {
    const x = left + index * 120;
    return `<g><rect x="${x}" y="2" width="10" height="10" rx="2" fill="${chartColor(index)}" /><text x="${x + 16}" y="11" font-size="12" fill="#475569">${escapeHtml(name)}</text></g>`;
  }).join('');

  return `<div class="block chart presentation-chart"${blockStyle(block)}>
    <div class="chart-title">${escapeHtml(title)}</div>
    <svg viewBox="0 0 ${svgWidth} ${svgHeight}" role="img" aria-label="${escapeHtml(title)}">
      <rect x="0" y="0" width="${svgWidth}" height="${svgHeight}" fill="#ffffff" />
      ${legend}
      ${grid}
      <line x1="${left}" y1="${top + plotHeight}" x2="${svgWidth - right}" y2="${top + plotHeight}" stroke="#94a3b8" stroke-width="1" />
      ${bars}
    </svg>
  </div>`;
}

function chartBarWidth(value: number, max: number): number {
  return Math.max(0, Math.min(100, (value / max) * 100));
}

function formatChartValue(value: number): string {
  return Number.isInteger(value) ? value.toString() : value.toFixed(2);
}

function chartColor(index: number): string {
  const colors = ['#2563eb', '#f97316', '#16a34a', '#9333ea', '#dc2626', '#0f766e'];
  return colors[index % colors.length];
}

function renderSourceBackedChartPlaceholder(block: MarkupBlock): string {
  const attributes = block.Attributes ?? {};
  const source = block.Source ?? chartAttr(attributes, 'source', 'range') ?? '';
  const sourceKind = chartSourceKind(source);
  const target = chartAttr(attributes, 'cell', 'target', 'anchor');
  const width = chartAttr(attributes, 'width', 'w');
  const height = chartAttr(attributes, 'height', 'h');
  const placement = target ? ` anchored at ${target}` : '';
  const size = width && height ? `, ${width} x ${height}` : '';

  return `<div class="chart-placeholder">Native editable ${sourceKind} chart from ${escapeHtml(source)}${escapeHtml(placement)}${escapeHtml(size)}.</div>`;
}

function renderChartMeta(block: MarkupBlock): string {
  const attributes = block.Attributes ?? {};
  const items: string[] = [];
  const source = block.Source ? (block.Sheet ? `${block.Sheet}!${block.Source}` : block.Source) : '';
  const target = chartAttr(attributes, 'cell', 'target', 'anchor');
  const width = chartAttr(attributes, 'width', 'w');
  const height = chartAttr(attributes, 'height', 'h');

  addChartMeta(items, 'type', block.ChartType);
  addChartMeta(items, 'source', source);
  addChartMeta(items, 'source kind', source ? chartSourceKind(source) : undefined);
  addChartMeta(items, 'target', target);
  addChartMeta(items, 'size', width && height ? `${width} x ${height}` : undefined);
  addChartMeta(items, 'x title', chartAttr(attributes, 'category-title', 'categoryTitle', 'x-title', 'xTitle', 'x-axis-title', 'xAxisTitle'));
  addChartMeta(items, 'y title', chartAttr(attributes, 'value-title', 'valueTitle', 'y-title', 'yTitle', 'y-axis-title', 'yAxisTitle'));
  addChartMeta(items, 'x format', chartAttr(attributes, 'category-format', 'categoryFormat', 'x-format', 'xFormat', 'category-number-format', 'categoryNumberFormat'));
  addChartMeta(items, 'y format', chartAttr(attributes, 'value-format', 'valueFormat', 'y-format', 'yFormat', 'value-number-format', 'valueNumberFormat'));
  addChartMeta(items, 'legend', chartAttr(attributes, 'legend', 'legend-position', 'legendPosition'));
  addChartMeta(items, 'labels', chartAttr(attributes, 'labels', 'data-labels', 'dataLabels'));
  addChartMeta(items, 'label position', chartAttr(attributes, 'label-position', 'labelPosition', 'data-label-position', 'dataLabelPosition'));
  addChartMeta(items, 'label format', chartAttr(attributes, 'label-format', 'labelFormat', 'data-label-format', 'dataLabelFormat'));
  addChartMeta(items, 'gridlines', chartAttr(attributes, 'gridlines', 'value-gridlines', 'valueGridlines', 'category-gridlines', 'categoryGridlines'));

  if (items.length === 0) {
    return '';
  }

  return `<div class="chart-meta">${items.join('')}</div>`;
}

function addChartMeta(items: string[], label: string, value?: string): void {
  if (!value || value.trim().length === 0) {
    return;
  }

  items.push(`<span class="chart-chip">${escapeHtml(label)}: ${escapeHtml(value)}</span>`);
}

function chartSourceKind(source: string): string {
  const localSource = source.includes('!') ? source.substring(source.lastIndexOf('!') + 1) : source;
  return /^\$?[A-Za-z]{1,3}\$?\d+\s*:\s*\$?[A-Za-z]{1,3}\$?\d+$/.test(localSource.trim())
    ? 'range'
    : 'table';
}

function chartAttr(attributes: Record<string, string>, ...names: string[]): string | undefined {
  for (const name of names) {
    const value = attributes[name];
    if (value && value.trim().length > 0) {
      return value;
    }
  }

  return undefined;
}

function blockStyle(block: MarkupBlock): string {
  const parts = styleParts(block);
  const position = block.Position;
  if (position && (position.X || position.Y || position.Width || position.Height)) {
    parts.push('position:absolute');
    if (position.X) {
      parts.push(`left:${cssLength(position.X)}`);
    }

    if (position.Y) {
      parts.push(`top:${cssLength(position.Y)}`);
    }

    if (position.Width) {
      parts.push(`width:${cssLength(position.Width)}`);
    }

    if (position.Height) {
      parts.push(`height:${cssLength(position.Height)}`);
    }
  }

  return parts.length === 0 ? '' : ` style="${parts.join(';')}"`;
}

function styleParts(block: MarkupBlock): string[] {
  const style = block.ResolvedStyle;
  if (!style) {
    return [];
  }

  const parts: string[] = [];
  const fontName = cssFontName(style.FontName);
  if (fontName) {
    parts.push(`font-family:${fontName}`);
  }

  if (typeof style.FontSize === 'number' && Number.isFinite(style.FontSize)) {
    parts.push(`font-size:${style.FontSize}pt`);
  }

  if (style.Bold !== undefined) {
    parts.push(`font-weight:${style.Bold ? 700 : 400}`);
  }

  if (style.Italic !== undefined) {
    parts.push(`font-style:${style.Italic ? 'italic' : 'normal'}`);
  }

  const textColor = cssColor(style.TextColor);
  if (textColor) {
    parts.push(`color:${textColor}`);
  }

  const fillColor = cssColor(style.FillColor);
  if (fillColor) {
    parts.push(`background:${fillColor}`);
  }

  const borderColor = cssColor(style.BorderColor);
  if (borderColor) {
    parts.push(`border-color:${borderColor}`);
  }

  const textAlign = cssTextAlign(style.TextAlign);
  if (textAlign) {
    parts.push(`text-align:${textAlign}`);
  }

  return parts;
}

function cssLength(value: string): string {
  const trimmed = value.trim();
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    return `${trimmed}in`;
  }

  return escapeHtml(trimmed);
}

function cssColor(value: string | undefined): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) {
    return undefined;
  }

  if (/^#?[0-9a-fA-F]{6}$/.test(trimmed)) {
    return trimmed.startsWith('#') ? trimmed : `#${trimmed}`;
  }

  return undefined;
}

function cssFontName(value: string | undefined): string | undefined {
  const trimmed = value?.replace(/[;"{}]/g, '').trim();
  return trimmed ? `"${trimmed}"` : undefined;
}

function cssTextAlign(value: string | undefined): string | undefined {
  const normalized = value?.trim().toLowerCase();
  return normalized === 'left' || normalized === 'center' || normalized === 'right' || normalized === 'justify'
    ? normalized
    : undefined;
}

function cssVerticalAlign(value: string | undefined): string | undefined {
  const normalized = value?.trim().toLowerCase();
  switch (normalized) {
    case 'top':
      return 'top';
    case 'middle':
    case 'center':
    case 'centre':
      return 'middle';
    case 'bottom':
      return 'bottom';
    default:
      return undefined;
  }
}

function cssBorderStyle(value: string | undefined): string | undefined {
  const normalized = value?.trim().toLowerCase();
  switch (normalized) {
    case 'true':
    case 'yes':
    case 'on':
    case '1':
    case 'thin':
    case 'hair':
      return 'solid';
    case 'medium':
    case 'thick':
    case 'double':
      return normalized;
    case 'dashed':
    case 'mediumdashed':
    case 'dashdot':
    case 'dashdotdot':
    case 'mediumdashdot':
    case 'mediumdashdotdot':
    case 'slantdashdot':
      return 'dashed';
    case 'dotted':
      return 'dotted';
    default:
      return undefined;
  }
}

function isExtension(block: MarkupBlock | undefined, command: string): boolean {
  return block?.Kind === 'Extension' && (block.Command ?? '').toLowerCase() === command;
}

function isColumns(block: MarkupBlock | undefined): boolean {
  return block?.Kind === 'Columns' || isExtension(block, 'columns');
}

function isColumn(block: MarkupBlock | undefined): boolean {
  const command = (block?.Command ?? '').toLowerCase();
  return block?.Kind === 'Column'
    || (block?.Kind === 'Extension' && (command === 'column' || command === 'left' || command === 'right'));
}

function columnBody(block: MarkupBlock): string {
  return block.Body ?? '';
}

function blocksFromMarkdownishBody(body: string): MarkupBlock[] {
  return body
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0)
    .map((line) => {
      const heading = /^(#{1,6})\s+(.*)$/.exec(line.trim());
      if (heading) {
        return { Kind: 'Heading', Text: heading[2] };
      }

      const item = /^[-*]\s+(.*)$/.exec(line.trim());
      if (item) {
        return { Kind: 'List', Items: [{ Text: item[1] }] };
      }

      return { Kind: 'Paragraph', Text: line.trim() };
    });
}

function renderMarkdownish(body: string, mode: PreviewRenderMode = 'document'): string {
  return blocksFromMarkdownishBody(body).map((block) => renderBlock(block, mode)).join('');
}

function renderInlineMarkdownish(body: string): string {
  return escapeHtml(body.trim()).replace(/\r?\n/g, '<br>');
}

function renderDiagnostic(item: MarkupDiagnostic): string {
  return `<div class="diagnostic">${escapeHtml(item.Severity ?? 'Info')}: ${escapeHtml(item.Message ?? '')}</div>`;
}

function renderPreviewActions(): string {
  return `<div class="toolbar-actions">
    <button class="toolbar-button" type="button" data-command="refresh">Refresh</button>
    <button class="toolbar-button" type="button" data-command="validate">Validate</button>
    <button class="toolbar-button" type="button" data-command="generateArtifacts">Generate Artifacts</button>
    <button class="toolbar-button" type="button" data-command="exportOfficeAndOpen">Export and Open</button>
    <button class="toolbar-button" type="button" data-command="openOutputFolder">Open Output Folder</button>
  </div>`;
}

function renderPreviewScript(scriptUri?: string): string {
  const mermaidScript = scriptUri ? `<script src="${escapeHtml(scriptUri)}"></script>` : '';
  return `${mermaidScript}<script>
    (function() {
      const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : undefined;
      if (!vscode) {
        return;
      }

      document.querySelectorAll('[data-command]').forEach((element) => {
        element.addEventListener('click', () => {
          const command = element.getAttribute('data-command');
          if (command) {
            vscode.postMessage({ command });
          }
        });
      });
    })();
  </script>`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeHtmlAttribute(value: string): string {
  return escapeHtml(value);
}

function previewOutputLabel(document: vscode.TextDocument, outputDirectory: vscode.Uri): string {
  const sourceDirectory = path.dirname(document.fileName);
  if (path.normalize(outputDirectory.fsPath) === path.normalize(sourceDirectory)) {
    return `Outputs default to source folder: ${outputDirectory.fsPath}`;
  }

  return `Outputs default to: ${outputDirectory.fsPath}`;
}
