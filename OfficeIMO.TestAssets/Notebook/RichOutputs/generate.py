"""Generate the saved chart outputs from the original figures below."""

import argparse
import base64
import hashlib
from pathlib import Path
from textwrap import dedent

import altair as alt
import nbformat
import plotly.io as pio
import vl_convert
from altair.utils.mimebundle import spec_to_mimebundle
from nbformat.v4 import new_code_cell, new_markdown_cell, new_notebook, new_output


SAMPLE = (
    Path(__file__).resolve().parents[3]
    / "Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/notebook-report.ipynb"
)


SETUP = """\
import altair as alt
import numpy as np
import plotly.graph_objects as go
from altair.utils.mimebundle import spec_to_mimebundle
from IPython.display import display
"""

PLOTLY_BAR = """\
chart = go.Figure(go.Bar(
    x=["Jan", "Feb", "Mar"],
    y=np.array([3, 12, 7], dtype=np.int32),
    name="Revenue", marker_color="#176c68",
))
chart.update_layout(title="Monthly revenue", xaxis_title="Month", yaxis_title="Revenue")
chart.show(renderer="plotly_mimetype")
"""

PLOTLY_SCATTER = """\
chart = go.Figure(go.Scatter(
    x=np.array([1, 3, 10], dtype=np.float64),
    y=np.array([4, 9, 5], dtype=np.float64),
    mode="markers", name="Observations", marker_color="#9b3d00",
))
chart.update_layout(title="Revenue observations", xaxis_title="Sample", yaxis_title="Revenue")
chart.show(renderer="plotly_mimetype")
"""

ALTAIR_BAR = """\
monthly_revenue = [
    {"month": "Jan", "revenue": 3},
    {"month": "Feb", "revenue": 12},
    {"month": "Mar", "revenue": 7},
]
chart = alt.Chart(alt.NamedData(name="monthly_revenue")).mark_bar(color="#254fb5").encode(
    x=alt.X("month:N", title="Month", sort=["Jan", "Feb", "Mar"]),
    y=alt.Y("revenue:Q", title="Revenue"),
).properties(title="Revenue by month", width=490, height=240,
             datasets={"monthly_revenue": monthly_revenue})
display(spec_to_mimebundle(chart.to_dict(), "vega-lite", mode="vega-lite", vegalite_version="6.4"), raw=True)
"""

PLOTLY_HEATMAP = """\
chart = go.Figure(go.Heatmap(
    x=["Jan", "Feb", "Mar"], y=["North", "Central", "South"],
    z=[[3, 12, 7], [8, 4, 13], [2, 11, 6]], colorscale="Viridis",
    texttemplate="%{z}", textfont={"size": 18},
))
chart.update_layout(
    title="Regional revenue heatmap", xaxis_title="Month", yaxis_title="Region",
    width=620, height=370, font={"family": "Arial"},
)
chart.show(renderer="plotly_mimetype+png")
"""

ALTAIR_LAYERED = """\
chart_data = alt.Data(values=[
    {"week": 1, "orders": 14}, {"week": 2, "orders": 19},
    {"week": 3, "orders": 15}, {"week": 4, "orders": 24},
    {"week": 5, "orders": 28},
])
base = alt.Chart(chart_data).encode(
    x=alt.X("week:Q", title="Week", scale=alt.Scale(zero=False)),
    y=alt.Y("orders:Q", title="Orders"),
)
chart = (base.mark_line(strokeWidth=4, color="#176c68")
         + base.mark_point(filled=True, size=140, color="#b54036")).properties(
    title="Weekly orders trend", width=515, height=225,
)
display(spec_to_mimebundle(chart.to_dict(), "vega-lite", mode="vega-lite", vegalite_version="6.4"), raw=True)
"""

TABLE = """\
from IPython.display import HTML
display(HTML("<table><thead><tr><th>Region</th><th>Revenue</th></tr></thead>"
             "<tbody><tr><td>North</td><td>22</td></tr>"
             "<tr><td>Central</td><td>25</td></tr>"
             "<tr><td>South</td><td>19</td></tr></tbody></table>"))
"""


def figure_from_source(source, namespace):
    # The final line is the notebook's display call. The generator captures its
    # MIME bundle directly below, without requiring a notebook server or kernel.
    chart_source = source.rstrip().rsplit("\n", 1)[0]
    exec(compile(chart_source, "<fixture chart source>", "exec"), namespace)
    return namespace["chart"]


def plotly_output(figure, with_png):
    renderer = pio.renderers["plotly_mimetype"]
    data = renderer.to_mimebundle(figure.to_dict())
    if with_png:
        png = pio.to_image(figure, format="png")
        data["image/png"] = base64.b64encode(png).decode("ascii")
    return new_output("display_data", data=data)


def altair_output(chart, with_images):
    spec = chart.to_dict()
    data = spec_to_mimebundle(spec, "vega-lite", mode="vega-lite", vegalite_version="6.4")
    if with_images:
        data["image/svg+xml"] = vl_convert.vegalite_to_svg(spec, vl_version="6.4")
        png = vl_convert.vegalite_to_png(spec, vl_version="6.4")
        data["image/png"] = base64.b64encode(png).decode("ascii")
    return new_output("display_data", data=data)


def summary_output():
    rows = [
        {"region": "North", "revenue": 22},
        {"region": "Central", "revenue": 25},
        {"region": "South", "revenue": 19},
    ]
    html_rows = "".join(
        f"<tr><td>{row['region']}</td><td>{row['revenue']}</td></tr>"
        for row in rows
    )
    return new_output("display_data", data={
        "application/vnd.dataresource+json": {
            "schema": {"fields": [
                {"name": "region", "type": "string"},
                {"name": "revenue", "type": "integer"},
            ]},
            "data": rows,
        },
        "text/html": (
            "<table><thead><tr><th>Region</th><th>Revenue</th></tr></thead>"
            f"<tbody>{html_rows}</tbody></table>"
        ),
    })


def build_notebook():
    import numpy as np
    import plotly.graph_objects as go

    namespace = {"alt": alt, "np": np, "go": go}
    notebook = new_notebook(
        metadata={
            "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
            "language_info": {"name": "python"},
        },
        cells=[
            new_markdown_cell(dedent("""\
                # Revenue and orders notebook

                This report uses sample results from three sales regions. Monthly
                revenue is $R_m = \\sum_r R_{m,r}$.
                """).rstrip(), id="intro"),
            new_code_cell(SETUP.rstrip(), id="setup", execution_count=1),
        ],
    )
    displays = [
        ("plotly-bar", PLOTLY_BAR, plotly_output, False),
        ("plotly-scatter", PLOTLY_SCATTER, plotly_output, False),
        ("altair-bar", ALTAIR_BAR, altair_output, False),
        ("plotly-heatmap", PLOTLY_HEATMAP, plotly_output, True),
        ("altair-layered", ALTAIR_LAYERED, altair_output, True),
    ]
    for execution_count, (cell_id, source, render, with_images) in enumerate(displays, start=2):
        figure = figure_from_source(source, namespace)
        notebook.cells.append(new_code_cell(
            source.rstrip(), id=cell_id, execution_count=execution_count,
            outputs=[render(figure, with_images)],
        ))
    notebook.cells.append(new_code_cell(
        TABLE.rstrip(), id="summary-table", execution_count=7, outputs=[summary_output()],
    ))
    nbformat.validate(notebook)
    return notebook


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=SAMPLE)
    args = parser.parse_args()
    notebook = build_notebook()
    with args.output.open("w", encoding="utf-8", newline="\n") as destination:
        nbformat.write(notebook, destination)
    print(f"{hashlib.sha256(args.output.read_bytes()).hexdigest()}  {args.output}")


if __name__ == "__main__":
    main()
