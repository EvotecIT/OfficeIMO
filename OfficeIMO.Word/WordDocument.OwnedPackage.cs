using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        // The editable package over _ownedPackageStream, when its layout allowed one; see WordZipPackage.
        private WordZipPackage? _ownedZipPackage;

        /// <summary>
        /// Opens an editable package held in <paramref name="packageStream"/>, which this document owns. Saving then
        /// rewrites only the parts that changed. Packages the in-memory package does not handle, and Strict documents,
        /// open through System.IO.Packaging as before.
        /// </summary>
        private static WordprocessingDocument OpenOwnedEditablePackage(MemoryStream packageStream, OpenSettings settings,
            out WordZipPackage? ownedZipPackage) {
            ownedZipPackage = WordZipPackage.TryOpen(packageStream);
            if (ownedZipPackage != null) {
                try {
                    WordprocessingDocument document = WordprocessingDocument.Open(ownedZipPackage, settings);
                    if (!document.StrictRelationshipFound) return document;
                } catch (Exception exception) when (exception is not OutOfMemoryException) {
                    // System.IO.Packaging decides: it opens the package or reports the problem as it always has.
                }
                // The SDK document is dropped without saving, and the package can no longer write to the stream.
                ownedZipPackage.Abandon();
                ownedZipPackage = null;
                packageStream.Position = 0;
            }
            return WordprocessingDocument.Open(packageStream, true, settings);
        }

        /// <summary>Whether saving to <paramref name="filePath"/> changes the document type or removes macros.</summary>
        private bool NeedsDocumentTypeAlignment(string filePath) {
            WordprocessingDocumentType documentType = GetDocumentType(filePath);
            return _wordprocessingDocument.DocumentType != documentType ||
                (!IsMacroEnabledDocumentType(documentType) && _wordprocessingDocument.MainDocumentPart?.VbaProjectPart != null);
        }
    }
}
