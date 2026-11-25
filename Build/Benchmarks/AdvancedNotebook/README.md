# Advanced notebook benchmarks

This opt-in PowerForge suite measures cold file export from a saved notebook to
a PDF or an interactive HTML file. Each measurement starts a fresh OfficeIMO
.NET or nbconvert Python process and ends when the process exits. The suite
checks the saved artifact outside timing. Notebook cells stay unexecuted.

- **PDF** uses the same saved Plotly MIME in both processes. The original set
  contains scatter3d, surface, and 2,300-point scattergl. `-ChartSet Expanded`
  uses a survey notebook with a triangular mesh, a geographic route, an inline
  GeoJSON choropleth, a heatmap, and mixed scatter/bar subplots. The nbconvert
  route takes a Plotly/Kaleido snapshot before calling WebPDF. Poppler checks
  the saved heading and visible colored chart. The OfficeIMO result also has
  to pass the vector-only check. Inspect the qualification PNGs
  for titles, axes, colors, and geometry. Layout, basemap details, tessellation,
  and shading can differ between the two static renderers.
- **Chart HTML** saves the same typed heatmap as an interactive offline HTML
  page. The nbconvert route uses the Plotly.py HTML renderer with its bundled
  Plotly.js staged locally before timing. HtmlTinkerX checks the exact reading
  `18.5` at `x=2`, `y=25` by hovering, then verifies zoom and reset in both
  outputs. It also verifies keyboard and pointer pan in the OfficeIMO view.
  The comparison times file export. Browser validation runs separately.
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

The PDF lane also requires Poppler's `pdftotext`, `pdftoppm`, and `pdfimages` on PATH. When
using nbconvert WebPDF, the Python environment owns its Playwright browser.

```powershell
./CompareAdvancedNotebook.ps1 -Mode Pdf -PythonPath $python -Chrome $chrome -QualifyOnly
./CompareAdvancedNotebook.ps1 -Mode Pdf -PythonPath $python -Chrome $chrome
```

For the expanded PDF set, save Plotly's world topology to a local directory
for Kaleido. The runner checks the SHA-256 of `world_110m.json` and gives
Kaleido a `file://` URI for offline access. The expected hash is
`d75915eaa31c870df6b972c9e5bb86910197825f33dcfef740f3b2f68cffe843`.

```powershell
$topojson = Join-Path ([IO.Path]::GetTempPath()) 'officeimo-plotly-oracle'
New-Item -ItemType Directory -Path $topojson -Force | Out-Null
Invoke-WebRequest https://cdn.plot.ly/world_110m.json -OutFile (Join-Path $topojson 'world_110m.json')
./CompareAdvancedNotebook.ps1 -Mode Pdf -ChartSet Expanded -PythonPath $python -Chrome $chrome -TopojsonDirectory $topojson
```

For Chart HTML, build the browser verifier and provision its local Chromium.
The package assets come from the pinned Plotly.py installation. The
`ArtifactBytes` metric covers the HTML file. nbconvert's `assets/plotly.min.js`
is a separate file and its size can be read from the qualification directory.

```powershell
dotnet build ./WidgetVerifier/WidgetVerifier.csproj -c Release
dotnet run --project ./WidgetVerifier/WidgetVerifier.csproj -c Release --no-build -- --install-browser
./CompareAdvancedNotebook.ps1 -Mode ChartHtml -PythonPath $python
```

To check the five-figure OfficeIMO viewer at wide and compact widths, export
the complete [survey notebook](../../../OfficeIMO.TestAssets/Notebook/AdvancedOutputs/Charts/plotly-interactive.ipynb)
with the driver and use the verifier's notebook mode. It checks navigation
and horizontal overflow at 1080px and 390px.

```powershell
dotnet run --project ./OfficeDriver/OfficeNotebookExport.csproj -c Release --no-build -- chart-html ../../../OfficeIMO.TestAssets/Notebook/AdvancedOutputs/Charts/plotly-interactive.ipynb survey.html
dotnet run --project ./WidgetVerifier/WidgetVerifier.csproj -c Release --no-build -- --chart-notebook survey.html wide.png compact.png
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
size reports HTML bytes. The external assets are separate files. Use `-Plan`
to resolve the cases and engines without saving samples.
