# Saved notebook chart fixture

The browser converter's sample also supplies the notebook conversion test
fixture. It contains original sample data and figures covered by the
[repository's MIT license](../../../LICENSE).

| Saved output | Producer representation | Visible content |
| --- | --- | --- |
| `plotly-bar` | Plotly JSON with an `i4` array | Jan, Feb and Mar revenue of 3, 12 and 7 |
| `plotly-scatter` | Plotly JSON with `f8` arrays | Three markers at (1, 4), (3, 9) and (10, 5) |
| `altair-bar` | Vega-Lite 6 JSON with a named inline dataset | Three month bars shown in calendar order |
| `plotly-heatmap` | Plotly JSON and Kaleido PNG (620 × 370) | Annotated 3 × 3 heatmap with a color scale |
| `altair-layered` | Vega-Lite 6 JSON, vl-convert SVG and PNG (565 × 290) | Five labeled weeks with a line and contrasting markers |
| `summary-table` | Data Resource JSON and HTML table | Three region totals |

The notebook also includes Python source and an inline LaTeX equation.
Its format is Jupyter 4.5. The Plotly MIME data comes from its
`plotly_mimetype` renderer. Altair's `spec_to_mimebundle` serializer emits the
`application/vnd.vegalite.v6+json` MIME bundle with its Vega-Lite 6.4.1
specification. [Plotly renderer guide](https://plotly.com/python/renderers/),
[Altair serializer source](https://github.com/vega/altair/blob/v6.3.0/altair/utils/mimebundle.py).

## Regenerate

Run the opt-in producer from the repository root with Python 3.14 and
Chrome for Testing. Set `BROWSER_PATH` to the Chrome executable if Kaleido
cannot discover it. The recorded runs used Python 3.14.3, Chrome for Testing
153.0.8010.12, macOS 27.0.1 on arm64, and the package versions in
[`requirements.txt`](requirements.txt). [Kaleido browser requirements](https://plotly.com/python/static-image-export/#chrome).

```sh
python3.14 -m venv /tmp/officeimo-rich-notebook
/tmp/officeimo-rich-notebook/bin/python -m pip install -r OfficeIMO.TestAssets/Notebook/RichOutputs/requirements.txt
/tmp/officeimo-rich-notebook/bin/python OfficeIMO.TestAssets/Notebook/RichOutputs/generate.py
```

The generator validates the notebook with nbformat before writing it. Two runs in
the recorded environment produced identical bytes. Image bytes can vary with
Chrome, operating system, and installed fonts.

| File | SHA-256 |
| --- | --- |
| [`generate.py`](generate.py) | `e6153f704d5211ac4db206fb93c2f76dea4aad2ce5d4e9ad1abb0034198b9d7f` |
| [`notebook-report.ipynb`](../../../Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/notebook-report.ipynb) | `ad7356d1b7307579736b070c57134d8a5be181d7bd164d13de5268e937c872e7` |
