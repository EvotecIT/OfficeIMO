# Saved Jupyter widget fixtures

The notebooks are byte-identical captures from their upstream repositories. They
contain saved views and model state. The notebook source serves as fixture data.

| Fixture | Upstream file | SHA-256 | License |
| --- | --- | --- | --- |
| [`widget-list.ipynb`](widget-list.ipynb) | [Jupyter nbconvert Widget List](https://github.com/jupyter/nbconvert/blob/f31a097873baba2216972969e59022cf98245cae/tests/files/Widget_List.ipynb) | `f19ad2075cc57b504ef2d83cc9f0f19cf025d108e880d1e5d8c9c6940501c9dd` | [BSD-3-Clause](LICENSE.nbconvert.txt) |
| [`anywidget-counter.ipynb`](anywidget-counter.ipynb) | [anywidget Counter](https://github.com/manzt/anywidget/blob/35160ac08e75918faee008a3b8996ab2a0a16a46/docs/public/notebooks/counter.ipynb) | `ad5a4c5509875ca66b5052aadfaf10c9ff03647c38cc55bf1420f474798cd920` | [MIT](LICENSE.anywidget.txt) |

Widget List has 41 resolved views and 187 saved models from older ipywidgets
versions. Its saved `HBox` contains a `Play` control and an `IntSlider` connected
through `LinkModel`. Both values start at 50. It also contains nested layouts,
standard form controls, a binary image, and views outside the OfficeIMO browser
control profile.

The anywidget example has three resolved views and four saved models. The
counter's custom `AnyModel` carries its renderer as executable JavaScript in
`_esm`. It exercises the HTML fallback for a custom widget.
