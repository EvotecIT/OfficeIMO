"""Run the original chart cells and save their Plotly.py notebook MIME outputs."""

from pathlib import Path
from textwrap import dedent

import nbformat
import plotly.io as pio


CELLS = {
    "survey-mesh": """
        import numpy as np
        import plotly.graph_objects as go

        # Elevations sampled at nine positions on a survey grid.
        east, north = np.meshgrid([0.0, 30.0, 60.0], [0.0, 30.0, 60.0])
        height = np.array([[3.0, 14.0, 5.0], [7.5, 24.0, 12.5], [4.0, 19.0, 6.0]])
        fig = go.Figure(go.Mesh3d(
            x=east.ravel(), y=north.ravel(), z=height.ravel(),
            i=np.array([0, 1, 1, 2, 3, 4, 4, 5], dtype=np.uint16),
            j=np.array([1, 4, 2, 5, 4, 7, 5, 8], dtype=np.uint16),
            k=np.array([3, 3, 4, 4, 6, 6, 7, 7], dtype=np.uint16),
            color="#237f86", flatshading=True, opacity=0.92,
        ))
        fig.update_layout(title_text="Surveyed ridge mesh", width=700, height=480)
        fig.update_scenes(xaxis_title="Easting", yaxis_title="Northing", zaxis_title="Elevation",
                          camera_eye=dict(x=1.7, y=1.25, z=1.45))
    """,
    "survey-route": """
        import numpy as np
        import plotly.graph_objects as go

        # The markers trace an illustrative sampling trip on the lower Hudson River.
        fig = go.Figure(go.Scattergeo(
            lon=np.array([-74.04, -74.03, -73.98, -73.93, -73.90]),
            lat=np.array([40.65, 40.71, 40.76, 40.83, 40.92]),
            mode="lines+markers", marker=dict(size=9, color="#8e356b"),
            line=dict(color="#8e356b", width=3),
        ))
        fig.update_layout(title_text="Hudson survey route", width=700, height=480)
        fig.update_geos(projection_type="equirectangular", lonaxis_range=[-74.40, -73.6],
                        lataxis_range=[40.55, 41.05])
    """,
    "survey-regions": """
        import numpy as np
        import plotly.graph_objects as go

        # Deliberately simplified sampling blocks around the lower Hudson River.
        def region(name, left, bottom, right, top):
            outline = [[left, bottom], [left, top], [right, top], [right, bottom], [left, bottom]]
            return {
                "type": "Feature", "id": name, "properties": {"station": name},
                "geometry": {"type": "Polygon", "coordinates": [outline]},
            }

        regions = {"type": "FeatureCollection", "features": [
            region("Upper reach", -74.00, 40.80, -73.88, 40.94),
            region("Middle reach", -74.07, 40.71, -73.95, 40.80),
            region("Lower reach", -74.14, 40.61, -74.02, 40.71),
        ]}
        fig = go.Figure(go.Choropleth(
            locations=["Upper reach", "Middle reach", "Lower reach"],
            z=np.array([12.0, 23.0, 34.0]), geojson=regions, colorscale="Viridis",
            zmin=0, zmax=40, colorbar=dict(title="Samples"),
            marker_line=dict(color="#374151", width=1),
        ))
        fig.update_layout(title_text="Hudson sampling regions", width=700, height=480)
        fig.update_geos(projection_type="equirectangular", fitbounds="locations",
                        showcoastlines=False, showland=False)
    """,
    "station-heatmap": """
        import numpy as np
        import plotly.graph_objects as go

        # Temperature observations by sampling hour and depth. One sensor stopped reporting.
        values = np.array([[13.0, 15.5, 16.0, 15.0], [16.5, 18.5, 19.5, 18.0],
                           [18.0, np.nan, 22.0, 20.5]], dtype=np.float64)
        fig = go.Figure(go.Heatmap(x=[1, 2, 3, 4], y=[10, 25, 40], z=values,
                                   zmin=10, zmax=24, colorscale="Viridis",
                                   colorbar=dict(title="Temperature")))
        fig.update_layout(title_text="Station temperature heatmap", width=700, height=480,
                          xaxis_title="Sampling hour", yaxis_title="Depth")
    """,
    "station-subplots": """
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots

        fig = make_subplots(rows=1, cols=2, subplot_titles=("Water temperature", "Samples per station"))
        fig.add_trace(go.Scatter(x=[1, 2, 3, 4], y=[14.5, 19.0, 21.0, 18.5],
                                 name="Surface", mode="lines+markers",
                                 line=dict(color="#285ca1")), row=1, col=1)
        fig.add_trace(go.Bar(x=["Upper", "Middle", "Lower"], y=[8, 12, 20], name="Day 1",
                             marker_color="#238263"), row=1, col=2)
        fig.add_trace(go.Bar(x=["Upper", "Middle", "Lower"], y=[10, 11, 18], name="Day 2",
                             marker_color="#b95525"), row=1, col=2)
        fig.update_xaxes(title_text="Sampling hour", row=1, col=1)
        fig.update_yaxes(title_text="Degrees Celsius", row=1, col=1)
        fig.update_xaxes(title_text="Station", row=1, col=2)
        fig.update_yaxes(title_text="Sample count", row=1, col=2)
        fig.update_layout(title_text="Field survey comparison", width=700, height=480, barmode="group")
    """,
}


def main():
    notebook = nbformat.v4.new_notebook(metadata={
        "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
        "language_info": {"name": "python"},
    })
    for execution_count, (identifier, source) in enumerate(CELLS.items(), 1):
        source = dedent(source).strip()
        variables = {}
        exec(compile(source, identifier, "exec"), variables)
        figure = variables["fig"]
        bundle = pio.renderers["plotly_mimetype"].to_mimebundle(figure.to_dict())
        assert list(bundle) == ["application/vnd.plotly.v1+json"]
        output = nbformat.v4.new_output(output_type="display_data", data=bundle, metadata={})
        notebook.cells.append(nbformat.v4.new_code_cell(
            source=source + "\nfig.show()", outputs=[output], id=identifier, execution_count=execution_count,
        ))
    nbformat.validate(notebook)
    nbformat.write(notebook, Path(__file__).with_name("plotly-interactive.ipynb"))


if __name__ == "__main__":
    main()
