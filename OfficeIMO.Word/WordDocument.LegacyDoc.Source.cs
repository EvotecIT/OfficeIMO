using OfficeIMO.Word.LegacyDoc.Write;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// The package the native DOC writer reads: the document's parts as they would be saved now. A document still on
        /// OfficeIMO's own Open XML layer is read in place, after the changes a load and save through the SDK path would
        /// make; one whose save has to go through the SDK opens it here. An SDK-backed document has already been prepared
        /// for saving by the caller.
        /// </summary>
        internal LegacyDocSourcePackage OpenLegacyDocSource() {
            if (_ownedModel is { } model) {
                if (WordOwnedSave.FindSaveBlocker(model) is string blocker) {
                    OpenSdkPackage("Save: " + blocker);
                } else {
                    WordOwnedSave.PrepareForSave(model);
                    return LegacyDocSourcePackage.FromModel(model);
                }
            }
            return LegacyDocSourcePackage.Open(CreateSavedPackageBytes(), PackageMaxCharactersInPart);
        }

        /// <summary>
        /// Writes a document whose SDK package was never opened as native DOC, when the save needs nothing the owned layer
        /// lacks. Returns false, having written nothing, when the save has to go through the SDK path.
        /// </summary>
        private bool TryCreateLegacyDocBytesWithoutSdk(WordSaveOptions? options, bool isTemplate, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            WordPackageModel? model = _ownedModel;
            // The owned layer holds only unsigned packages, so the signature policy has nothing to act on.
            if (model == null || ResolveSavePassword(options) != null || AutoUpdateToc) return false;
            if (WordOwnedSave.FindSaveBlocker(model) is string blocker) {
                OpenSdkPackage("Save: " + blocker);
                return false;
            }
            bytes = LegacyDocWriter.WriteDocument(this, options, isTemplate);
            return true;
        }

        /// <summary>
        /// The owned-layer counterpart of saving to a .doc or .dot path: checks the destination first, as the SDK path
        /// does before encoding. Returns false for other paths, password-protected output and saves that need the SDK.
        /// </summary>
        private bool TryCreateLegacyDocFileBytesWithoutSdk(string filePath, WordSaveOptions? options, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            if (_ownedModel == null || !IsLegacyDocPath(filePath) || ResolveSavePassword(options) != null) return false;
            EnsureDestinationFileWritable(filePath);
            return TryCreateLegacyDocBytesWithoutSdk(options, IsLegacyDocTemplatePath(filePath), out bytes);
        }

        private static bool IsLegacyDocTemplatePath(string filePath) =>
            string.Equals(Path.GetExtension(filePath), ".dot", StringComparison.OrdinalIgnoreCase);
    }
}
