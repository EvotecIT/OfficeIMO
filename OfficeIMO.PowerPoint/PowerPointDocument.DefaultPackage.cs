using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        // Cache only serialized bytes. Every creation opens its own expandable stream and mutable part tree.
        private static readonly Lazy<byte[]> DefaultPackage = new(() => CreateDefaultPackageBytes(applyTheme: false));
        private static readonly Lazy<byte[]> DefaultThemedPackage = new(() => CreateDefaultPackageBytes(applyTheme: true));

        /// <summary>Serializes the existing native scaffold once, with an optional canonical authoring theme.</summary>
        private static byte[] CreateDefaultPackageBytes(bool applyTheme) {
            using var packageStream = new MemoryStream(StreamBufferSize);
            using (PresentationDocument document = PresentationDocument.Create(
                       packageStream, PresentationDocumentType.Presentation, autoSave: false)) {
                var presentation = new PowerPointDocument(document, string.Empty, isNewDocument: true);
                if (applyTheme) presentation.ApplyLayoutTheme(OfficeTheme.Default);
                presentation.PresentationRoot.Save();
                document.Save();
            }
            return packageStream.ToArray();
        }
    }
}
