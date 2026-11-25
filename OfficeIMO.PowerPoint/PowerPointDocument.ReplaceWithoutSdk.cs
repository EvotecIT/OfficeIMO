using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Text replacement on a presentation whose SDK package is not open yet edits the slides' XML on the Open XML layer.
    /// The edits stay in a package the document holds until something reads the package: a save writes them with the
    /// document properties, and opening the SDK writes them into the bytes it opens.
    /// </summary>
    public sealed partial class PowerPointDocument {
        // Slide edits not yet written into the package bytes, and whether any tree in it has changed.
        private OfficeOpcEditablePackage? _pendingPackage;
        private bool _pendingEdits;

        /// <summary>Replaces text without the SDK, or returns false, having changed nothing, for the model to replace it.</summary>
        private bool TryReplaceAllWithoutSdk(string text, string replacement, PowerPointTextSearchOptions? options, out int count) {
            count = 0;
            // Notes are rewritten whole by the model.
            if (options?.IncludeNotes == true || GetPendingPackage() is not { } package) return false;
            try {
                List<string>? slides = ReadSlideParts(package);
                if (slides == null) return false;
                var search = new PowerPointTreeReplace.Search(text, replacement, OfficeTextSearch.GetComparison(options),
                    options?.WholeWords == true, options?.IncludeTables ?? true);
                // Every slide is planned before any is changed, so a slide that declines leaves the others as they were.
                var edits = new List<PowerPointTreeReplace.Edit>();
                int total = 0;
                foreach (string slide in slides) {
                    int found = PowerPointTreeReplace.PlanSlide(package.GetXml(slide), search, edits);
                    if (found < 0) return false;
                    total += found;
                }
                foreach (PowerPointTreeReplace.Edit edit in edits) edit.Tree.SetInnerText(edit.Text, edit.Value);
                _pendingEdits |= edits.Count > 0;
                count = total;
                return true;
            } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
                // The model reads the package with its own errors.
                return false;
            }
        }

        /// <summary>
        /// The package that holds edits made without the SDK, opened on first use; null when the document can't be edited
        /// that way: it is closed, read-only, has the SDK open, or the SDK has a part size limit to enforce.
        /// </summary>
        private OfficeOpcEditablePackage? GetPendingPackage() {
            if (_disposed || !_sdkDeferred || _sdkDocument != null || _accessMode != DocumentAccessMode.ReadWrite
                || _deferredOpenSettings?.MaxCharactersInPart > 0
                || _packageStream is not { CanRead: true, CanSeek: true } stream) {
                return null;
            }
            if (_pendingPackage == null) {
                stream.Position = 0;
                _pendingPackage = OfficeOpcEditablePackage.Open(stream, leaveOpen: true);
            }
            return _pendingPackage;
        }

        // The slide parts of the first slide list, in order; null when one is missing, external or listed twice.
        private static List<string>? ReadSlideParts(OfficeOpcEditablePackage package) =>
            PackageSlideList.Read(package)?.Entries.Select(entry => entry.Part).ToList();

        /// <summary>Writes pending slide edits into the package bytes, so whatever reads them next sees the edits.</summary>
        private void CommitPendingEdits() {
            OfficeOpcEditablePackage? pending = _pendingPackage;
            if (pending == null) return;
            _pendingPackage = null;
            _pendingRelationships = null;
            bool edited = _pendingEdits;
            _pendingEdits = false;
            if (!edited) {
                pending.Dispose();
                return;
            }
            byte[] bytes;
            using (pending) {
                using var output = new MemoryStream();
                pending.Save(output);
                bytes = output.ToArray();
            }
            Stream? previous = _packageStream;
            _packageStream = OfficePackageLoadSource.CreatePackageStream(bytes, _accessMode);
            previous?.Dispose();
        }

        /// <summary>Drops pending edits without writing them, when the document closes.</summary>
        private void DiscardPendingEdits() {
            _pendingPackage?.Dispose();
            _pendingPackage = null;
            _pendingRelationships = null;
            _pendingEdits = false;
        }
    }
}
