import { writeFileSync } from 'node:fs';
import { serializeNotebookToString } from './serializers.mts';

const notebook = {
  metadata: { nbformat: 4, nbformat_minor: 5, metadata: { language_info: { name: 'python' } } },
  cells: [
    { kind: 1, languageId: 'markdown', value: '# OfficeIMO VS Code evidence', metadata: { id: 'heading' } },
    { kind: 2, languageId: 'powershell', value: 'Write-Output "Saved from VS Code"', metadata: { id: 'powershell', execution_count: 1 },
      outputs: [{ items: [{ mime: 'application/vnd.code.notebook.stdout', data: new TextEncoder().encode('Saved from VS Code\n') }] }] }
  ]
};
writeFileSync('vscode.ipynb', serializeNotebookToString(notebook));
