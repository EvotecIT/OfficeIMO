"""Run nbconvert once from saved notebook state. Never execute notebook cells."""
import argparse
import base64
import os
from pathlib import Path

from nbconvert.exporters import HTMLExporter, WebPDFExporter
from nbconvert.preprocessors import Preprocessor

PLOTLY_MIME = "application/vnd.plotly.v1+json"


class PlotlySnapshots(Preprocessor):
    """Replace each supported saved Plotly MIME figure with its Kaleido PNG."""

    def preprocess_cell(self, cell, resources, index):
        import plotly.io

        for output in cell.get("outputs", []):
            data = output.get("data", {})
            spec = data.get(PLOTLY_MIME)
            if spec is None:
                continue
            layout = spec.get("layout", {})
            image = plotly.io.to_image(
                spec,
                format="png",
                width=layout.get("width", 700),
                height=layout.get("height", 480),
                validate=False,
            )
            data.clear()
            data["image/png"] = base64.b64encode(image).decode("ascii")
        return cell, resources


class PlotlyInteractive(Preprocessor):
    """Use Plotly's official HTML renderer with a bundled, pre-staged local runtime."""

    def preprocess_cell(self, cell, resources, index):
        from plotly.io import to_html

        for output in cell.get("outputs", []):
            data = output.get("data", {})
            spec = data.get(PLOTLY_MIME)
            if spec is None:
                continue
            html = to_html(spec, full_html=False, include_plotlyjs="assets/plotly.min.js",
                           include_mathjax=False, validate=False, div_id=f"saved-chart-{index}")
            data.clear()
            data["text/html"] = html
        return cell, resources


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=("Pdf", "WidgetHtml", "ChartHtml"))
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--chrome", type=Path)
    parser.add_argument("--topojson", type=Path)
    args = parser.parse_args()

    if args.mode == "Pdf":
        if args.chrome is None:
            parser.error("Pdf requires a local Chrome executable.")
        os.environ["BROWSER_PATH"] = str(args.chrome.resolve())
        import numpy
        import plotly
        import plotly.io
        from importlib.metadata import version

        if (plotly.__version__, version("kaleido"), numpy.__version__) != ("6.9.0", "1.5.2", "2.4.6"):
            parser.error("Use the pinned Plotly, Kaleido, and NumPy versions in requirements.txt.")
        if args.topojson is not None:
            plotly.io.defaults.topojson = args.topojson.resolve().as_uri() + "/"

        exporter = WebPDFExporter(
            allow_chromium_download=False,
            require_js_url="data:text/javascript,",
            mathjax_url="data:text/javascript,",
            exclude_input=True,
        )
        exporter.register_preprocessor(PlotlySnapshots(), enabled=True)
    elif args.mode == "WidgetHtml":
        exporter = HTMLExporter(
            require_js_url="assets/require.js",
            widget_renderer_url="assets/embed-amd.js",
            mathjax_url="data:text/javascript,",
            exclude_input=True,
        )
    else:
        exporter = HTMLExporter(
            require_js_url="data:text/javascript,",
            mathjax_url="data:text/javascript,",
            exclude_input=True,
        )
        exporter.register_preprocessor(PlotlyInteractive(), enabled=True)

    result, _ = exporter.from_filename(str(args.source))
    if isinstance(result, bytes):
        args.output.write_bytes(result)
    else:
        args.output.write_text(result, encoding="utf-8")


if __name__ == "__main__":
    main()
