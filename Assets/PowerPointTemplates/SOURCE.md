# PowerPoint source-producer corpus

The `.pptx` files in this folder are sanitized Microsoft PowerPoint-authored
fixtures. Their package metadata identifies `Microsoft Office PowerPoint` with
application version `16.0000`. `corpus-manifest.json` pins each fixture by
SHA-256 and records the contract it covers.

The corpus is immutable test input. Tests open each source, perform an edit in
memory, save to a new package, reopen that result, and validate the edited
package. Do not replace a fixture silently: update the manifest hash and this
provenance note in the same reviewed change, and explain which producer and
contract the replacement represents.
