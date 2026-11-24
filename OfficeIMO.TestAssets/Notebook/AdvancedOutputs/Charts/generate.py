"""Regenerate the JSON-only saved Plotly figures with the actual notebook MIME renderer."""

from pathlib import Path

import nbformat
import numpy as np
import plotly.graph_objects as go
import plotly.io as pio

# The helix coordinates are adapted from the MIT-licensed Plotly.py 3D scatter notebook.
theta = np.linspace(0, 10, 50)
axis = np.arange(-3, 4, dtype=np.float64)
xx, yy = np.meshgrid(axis, axis)
random = np.random.default_rng(42)

figures = (
    ("upstream-helix", "# Basic 3D Scatter Plot from the Plotly.py documentation",
     go.Figure(go.Scatter3d(x=np.cos(theta), y=np.sin(theta), z=theta, mode="markers")),
     "Helix observations"),
    ("surface", "# Cividis surface with a custom orthographic camera",
     go.Figure(go.Surface(x=axis, y=axis, z=np.sin(np.hypot(xx, yy)) * 2, colorscale="Cividis")),
     "Radial surface"),
    ("scattergl", "# A seeded two-dimensional WebGL scatter",
     go.Figure(go.Scattergl(x=random.normal(size=2300), y=random.normal(size=2300), mode="markers",
                            marker=dict(size=7, color="#b34636"))),
     "WebGL observations"),
)
notebook = nbformat.v4.new_notebook(metadata={"language_info": {"name": "python"}})
for identifier, source, figure, title in figures:
    figure.update_layout(title_text=title, width=700, height=480)
    if identifier == "surface":
        figure.update_scenes(camera_eye=dict(x=1.7, y=1.1, z=1.7),
                             camera_projection_type="orthographic",
                             xaxis_title="Distance east", yaxis_title="Distance north", zaxis_title="Height")
    bundle = pio.renderers["plotly_mimetype"].to_mimebundle(figure.to_dict())
    assert list(bundle) == ["application/vnd.plotly.v1+json"]
    output = nbformat.v4.new_output(output_type="display_data", data=bundle, metadata={})
    notebook.cells.append(nbformat.v4.new_code_cell(source=source, outputs=[output], id=identifier))
nbformat.validate(notebook)
nbformat.write(notebook, Path(__file__).with_name("plotly-advanced.ipynb"))
