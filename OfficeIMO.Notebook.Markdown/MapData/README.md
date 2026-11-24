# Offline geographic data

Made with Natural Earth. The embedded 1:110m land and country polygons are
intended for world and regional chart context at this scale. The packaged data
is public domain. [Natural Earth terms](https://www.naturalearthdata.com/about/terms-of-use/)

`NaturalEarth110m.json.gz` comes from Natural Earth Vector 5.1.2 at revision
`f1890d9f152c896d250a77557a5751a93d494776`. Its JSON lists country ISO codes,
English names, label longitude and latitude, and polygons. Its land list
contains polygons. Each polygon contains an exterior ring followed by any
holes. Source coordinates and label positions were rounded to four decimal
places, all other properties were dropped, and multipart geometry was kept in
source order. The JSON uses UTF-8 and compact separators, and the gzip header
has no filename and a zero timestamp.

| Source | SHA-256 |
| --- | --- |
| [Country polygons](https://github.com/nvkelso/natural-earth-vector/blob/f1890d9f152c896d250a77557a5751a93d494776/geojson/ne_110m_admin_0_countries.geojson) | `6866c877d39cba9c357620878839b336d569f8c662d3cfab4cb1dbe2d39c977f` |
| [Land polygons](https://github.com/nvkelso/natural-earth-vector/blob/f1890d9f152c896d250a77557a5751a93d494776/geojson/ne_110m_land.geojson) | `9e0729ee253ca7d7a5c4ae9395fb1902264c5377c52e224d13dd85010e2835d9` |
| Embedded asset | `a442b4d526139b94416bdb946a63cb0dd3863aa1d1865ea7aefea638666931ec` |

Country fields in the source are `ISO_A3_EH`, `NAME_EN`, `LABEL_X`, and
`LABEL_Y`.
