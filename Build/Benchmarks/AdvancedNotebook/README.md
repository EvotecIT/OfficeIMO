# Advanced notebook benchmarks

This opt-in PowerForge suite measures cold file export from a saved notebook to
a PDF or an interactive HTML file. Each measurement starts a fresh OfficeIMO
.NET or nbconvert Python process and ends when the process exits. The suite
checks the saved artifact outside timing. Notebook cells stay unexecuted.

- **PDF** uses the same saved Plotly MIME in both processes. The three case
  notebooks preserve the producer's scatter3d, surface, and 2,300-point scattergl
  output. The nbconvert route takes a Plotly/Kaleido snapshot before it calls
  WebPDF. Poppler checks the saved heading and visible colored chart. Inspect
  the qualification PNGs for titles, axes, colors, and geometry. The two
  static renderers can differ in layout, tessellation, and shading.
- **Widget HTML** uses the same selection from the upstream nbconvert Widget
  List fixture. HtmlTinkerX opens the saved result with external requests
  blocked, checks the initial slider value, changes it from the keyboard,
  and observes Play move the linked slider. The comparison times file export.
  The browser validation runs separately. nbconvert's local runtime assets
  are staged before timing.

## Prerequisites

Install PowerShell 7, PSPublishModule 3.0.194 or later, and the .NET SDK selected
by the repository. Create a Python environment with [the pinned requirements](requirements.txt).
Run the examples below from this directory. Set `$python` to your Python
environment executable and `$chrome` to a compatible, locally installed
Chrome or Chromium binary for Kaleido.

```powershell
& $python -m pip install -r requirements.txt
& $python -m playwright install chromium
dotnet build ./OfficeDriver/OfficeNotebookExport.csproj -c Release
```

The PDF lane also requires Poppler's `pdftotext` and `pdftoppm` on PATH. When
using nbconvert WebPDF, the Python environment owns its Playwright browser.

```powershell
./CompareAdvancedNotebook.ps1 -Mode Pdf -PythonPath $python -Chrome $chrome -QualifyOnly
./CompareAdvancedNotebook.ps1 -Mode Pdf -PythonPath $python -Chrome $chrome
```

For Widget HTML, extract the npm packages `@jupyter-widgets/html-manager@1.0.15`
and `requirejs@2.3.8`. Set `$widgetDist` to the first package's `dist`
directory and `$requireJs` to the second package's `require.js` file. Keep the
HTML manager's font files alongside `embed-amd.js`.

```powershell
dotnet build ./WidgetVerifier/WidgetVerifier.csproj -c Release
dotnet run --project ./WidgetVerifier/WidgetVerifier.csproj -c Release --no-build -- --install-browser
./CompareAdvancedNotebook.ps1 -Mode WidgetHtml -PythonPath $python -WidgetManagerDist $widgetDist -RequireJs $requireJs
```

The runner writes PDF or HTML qualification files, PNG previews, and per-sample
results under `Ignore/Benchmarks/AdvancedNotebook` by default. A sample enters
the comparison only after its artifact passes validation. The widget output
size reports HTML bytes. The external assets are separate files.
