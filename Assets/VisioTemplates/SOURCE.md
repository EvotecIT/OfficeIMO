# Visio source-producer corpus

The `VisioAdvancedRoadmap` packages were authored through the Microsoft Visio
16 COM object model and saved by Visio in each supported Open XML package
family. They contain a dense connected process, container and lane metadata,
nested group geometry, shape data, and ShapeSheet formulas.

The fixtures contain no VBA project. The macro-enabled variants prove the
package family and desktop open/save boundary; separate constructed fixtures
cover opaque VBA payload and relationship preservation without distributing a
macro-enabled binary.

`corpus-manifest.json` pins every fixture by SHA-256. Personal author metadata
was removed before the files were committed. Replace a fixture only with a
reviewed producer-generated package, and update the manifest and this note in
the same change.
