using System;
using System.IO.Packaging;
using System.IO;

namespace OfficeIMO.Word {
    public static partial class Helpers {
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

        public static void MakeOpenOfficeCompatible(Stream fileStream) {
            using (Package package = Package.Open(fileStream, FileMode.Open, FileAccess.ReadWrite)) {
                // Fix relationships in /_rels/.rels
                Uri globalRelsUri = new Uri("/_rels/.rels", UriKind.Relative);
                FixPartIfExists(package, globalRelsUri, false);

                // Fix relationships in /word/_rels/document.xml.rels (remove /word/)
                Uri documentRelsUri = new Uri("/word/_rels/document.xml.rels", UriKind.Relative);
                FixPartIfExists(package, documentRelsUri, true);
            }
        }

        private static void FixPartIfExists(Package package, Uri partUri, bool removeWordPrefix) {
            if (package.PartExists(partUri)) {
                PackagePart part = package.GetPart(partUri);
                FixRelationships(part, removeWordPrefix);
            }
        }

        private static void FixRelationships(PackagePart relsPart, bool removeWordPrefix) {
            using (Stream stream = relsPart.GetStream(FileMode.Open, FileAccess.ReadWrite)) {
                var xml = System.Xml.Linq.XDocument.Load(stream);
                var relationships = xml.Root.Elements();

                bool modified = false;
                foreach (var relationship in relationships) {
                    var targetAttribute = relationship.Attribute("Target");
                    if (targetAttribute != null) {
                        if (removeWordPrefix && targetAttribute.Value.StartsWith("/word/")) {
                            targetAttribute.Value = targetAttribute.Value.Substring(6); // Remove "/word/"
                            modified = true;
                        } else if (!removeWordPrefix && targetAttribute.Value.StartsWith("/")) {
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
    }
}
