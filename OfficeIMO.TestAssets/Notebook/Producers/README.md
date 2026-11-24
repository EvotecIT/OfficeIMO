# Jupyter producer fixtures

`fixtures.json` pins the source and digest for each file. These files are inputs
for JSON loading, editing, and readback. The test suite never executes their
cells. The JupyterLab file comes from its integration tests. The Kaggle file
comes from Deepnote's Kaggle compatibility suite. The Databricks fixture
retains actual workspace export metadata from a public Apache-licensed
notebook.

The published files are byte-identical. The upstream notices are adjacent for
JupyterLab, Google Colab, and the Databricks export. The [Deepnote license](../Vendors/deepnote-LICENSE)
covers the Kaggle compatibility fixture and the original Deepnote snapshot.
The [Microsoft license](../Scripts/LICENSE.Polyglot.txt) covers the .NET
Interactive sample.

Three files were generated with pinned producer software.

- **VS Code.** Node 24.18.0 ran the upstream `serializeNotebookToString`
  from the pinned revision in the manifest. The driver passes a Markdown cell,
  a PowerShell code cell, and saved stdout through that serializer.
  For Node's TypeScript loader, the upstream `common.ts`, `constants.ts` and
  `serializers.ts` use `.mts` extensions, and `CellOutputMetadata` is marked
  as a type-only import.
- **Deepnote.** `@deepnote/convert@5.0.0` exported the unchanged
  [Hello World snapshot](../Vendors/deepnote-hello.snapshot.deepnote) with
  `deepnote-convert … --output-format jupyter -o output`. The fixture contains
  the single file written inside that directory.
- **marimo.** `marimo==0.25.2` and `nbformat==5.11.1` ran
  `marimo export ipynb marimo-source.py --include-outputs -o marimo.ipynb`.
  The checked-in source contains the two original cells used for generation.
