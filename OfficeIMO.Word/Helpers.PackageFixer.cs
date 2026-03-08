using System.IO.Packaging;

namespace OfficeIMO.Word {
    /// <summary>
    /// Helper methods for manipulating Open XML packages.
    /// </summary>
    public static partial class Helpers {
        /// <summary>
        /// Adjusts relationship targets so that the document can be opened by OpenOffice.
        /// </summary>
        /// <param name="filePath">Path to the document package.</param>
        public static void MakeOpenOfficeCompatible(string filePath) {
            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite)) {
                // Fix relationships in /_rels/.rels
                Uri globalRelsUri = new Uri("/_rels/.rels", UriKind.Relative);
                FixPartIfExists(package, globalRelsUri, false);

                // Fix relationships in /word/_rels/document.xml.rels (remove /word/)
                Uri documentRelsUri = new Uri("/word/_rels/document.xml.rels", UriKind.Relative);
                FixPartIfExists(package, documentRelsUri, true);
            }
        }

        /// <summary>
        /// Adjusts relationship targets for a document provided as a stream.
        /// </summary>
        /// <param name="fileStream">Stream containing the document package.</param>
        public static void MakeOpenOfficeCompatible(Stream fileStream) {
            using var nonDisposingStream = new NonDisposingStream(fileStream);
            using (Package package = Package.Open(nonDisposingStream, FileMode.Open, FileAccess.ReadWrite)) {
                // Fix relationships in /_rels/.rels
                Uri globalRelsUri = new Uri("/_rels/.rels", UriKind.Relative);
                FixPartIfExists(package, globalRelsUri, false);

                // Fix relationships in /word/_rels/document.xml.rels (remove /word/)
                Uri documentRelsUri = new Uri("/word/_rels/document.xml.rels", UriKind.Relative);
                FixPartIfExists(package, documentRelsUri, true);
            }
        }

        /// <summary>
        /// Updates relationships for the specified part if it exists in the package.
        /// </summary>
        /// <param name="package">Package to modify.</param>
        /// <param name="partUri">URI of the relationships part.</param>
        /// <param name="removeWordPrefix">Whether to remove the /word prefix from targets.</param>
        private static void FixPartIfExists(Package package, Uri partUri, bool removeWordPrefix) {
            if (package.PartExists(partUri)) {
                PackagePart part = package.GetPart(partUri);
                FixRelationships(part, removeWordPrefix);
            }
        }

        /// <summary>
        /// Fixes relationship targets in the specified relationships part.
        /// </summary>
        /// <param name="relsPart">Relationship part to process.</param>
        /// <param name="removeWordPrefix">Whether to remove the /word prefix from targets.</param>
        private static void FixRelationships(PackagePart relsPart, bool removeWordPrefix) {
            using (Stream stream = relsPart.GetStream(FileMode.Open, FileAccess.ReadWrite)) {
                var xml = System.Xml.Linq.XDocument.Load(stream);
                var root = xml.Root;
                if (root == null) return;
                var relationships = root.Elements();

                bool modified = false;
                foreach (var relationship in relationships) {
                    var targetAttribute = relationship.Attribute("Target");
                    if (targetAttribute != null) {
                        if (removeWordPrefix && targetAttribute.Value.StartsWith("/word/", StringComparison.Ordinal)) {
                            targetAttribute.Value = targetAttribute.Value.Substring(6); // Remove "/word/"
                            modified = true;
                        } else if (!removeWordPrefix && targetAttribute.Value.StartsWith("/", StringComparison.Ordinal)) {
                            targetAttribute.Value = targetAttribute.Value.TrimStart('/');
                            modified = true;
                        }
                    }
                }

                if (modified) {
                    stream.SetLength(0); // Clear the original content
                    xml.Save(stream); // Save the modified content
                }
            }
        }

        /// <summary>
        /// Wraps a stream while preventing disposal from closing the underlying instance.
        /// </summary>
        private sealed class NonDisposingStream : Stream {
            private readonly Stream _inner;

            public NonDisposingStream(Stream inner) {
                _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            }

            public override bool CanRead => _inner.CanRead;
            public override bool CanSeek => _inner.CanSeek;
            public override bool CanWrite => _inner.CanWrite;
            public override long Length => _inner.Length;

            public override long Position {
                get => _inner.Position;
                set => _inner.Position = value;
            }

            public override void Flush() => _inner.Flush();

            public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

            public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

            public override void SetLength(long value) => _inner.SetLength(value);

            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

            protected override void Dispose(bool disposing) {
                if (disposing && _inner.CanWrite) {
                    _inner.Flush();
                }
            }
        }
    }
}
