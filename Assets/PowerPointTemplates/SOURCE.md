# PowerPoint source-producer corpus

The `.pptx` files in this folder are sanitized Microsoft PowerPoint-authored
fixtures. Their package metadata identifies `Microsoft Office PowerPoint` with
application version `16.0000`. `corpus-manifest.json` pins each fixture by
SHA-256 and records the contract it covers.

`PowerPointAdvancedRoadmap.pptx` was authored through the PowerPoint 16 COM
object model. It contains native 3-D bar, 3-D line, 3-D area, 3-D pie,
Pie-of-Pie, stock, surface, and 3-D surface charts, plus PowerPoint-authored
SmartArt, animation timing, and embedded audio. Personal author and temporary
path metadata was removed before the fixture was committed. Slide titles are
frontmost so PowerPoint's native image exporter cannot composite opaque chart
surfaces over the title text.

The corpus is immutable test input. Tests open each source, perform an edit in
memory, save to a new package, reopen that result, and validate the edited
package. Do not replace a fixture silently: update the manifest hash and this
provenance note in the same reviewed change, and explain which producer and
contract the replacement represents.
