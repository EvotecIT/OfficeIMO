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
            if (options?.IncludeNotes == true || GetPackage(edit: true) is not { } package) return false;
            try {
                List<string>? slides = ReadSlideParts(package);
                if (slides == null) return false;
                var search = new PowerPointTreeReplace.Search(text, replacement, OfficeTextSearch.GetComparison(options),
                    options?.WholeWords == true, options?.IncludeTables ?? true);
                // Every slide is planned before any is changed, so a slide that declines leaves the others as they were.
                var edits = new List<PowerPointTreeReplace.Edit>();
                int total = 0;
                foreach (string slide in slides) {
                    // A slide over the part size limit is read by the model, which reports it.
                    if (IsOverPartLimit(package, slide)) return false;
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

        // The slide parts of the first slide list, in order; null when one is missing, external or listed twice.
        private static List<string>? ReadSlideParts(OfficeOpcEditablePackage package) =>
            PackageSlideList.Read(package)?.Entries.Select(entry => entry.Part).ToList();

        /// <summary>
        /// Writes pending edits into the package bytes, so whatever reads them next sees them. With
        /// <paramref name="keepPackage"/> the package stays open, so slides and shapes read from it keep working;
        /// otherwise it is closed, as before the SDK opens the bytes.
        /// </summary>
        private void CommitPendingEdits(bool keepPackage = false) {
            OfficeOpcEditablePackage? pending = _pendingPackage;
            if (pending == null) return;
            if (_pendingEdits) {
                using var output = new MemoryStream();
                pending.Save(output);
                ReplacePackageStream(OfficePackageLoadSource.CreatePackageStream(output.ToArray(), _accessMode));
                // Cleared only once the edited bytes are in place, so a failed write leaves the edits pending.
                _pendingEdits = false;
            }
            if (!keepPackage) ReleasePendingPackage();
        }

        /// <summary>Drops pending edits without writing them, when the document closes.</summary>
        private void DiscardPendingEdits() {
            _pendingEdits = false;
            ReleasePendingPackage();
        }

        // The stream the pending package was opened on, which it reads parts it has not parsed from; it stays open
        // while the package does, even after the document moved on to newer bytes.
        private Stream? _pendingSource;

        /// <summary>Makes <paramref name="next"/> the document's package bytes, closing the previous stream unless the
        /// pending package still reads from it.</summary>
        private void ReplacePackageStream(Stream next) {
            // The package the load checked reads the previous bytes.
            ReleaseLoadedPackage();
            Stream? previous = _packageStream;
            _packageStream = next;
            if (previous != null && !ReferenceEquals(previous, _pendingSource)) previous.Dispose();
        }

        private void ReleasePendingPackage() {
            _pendingPackage?.Dispose();
            _pendingPackage = null;
            _pendingRelationships = null;
            _classicCommentReadPackage = null;
            _classicCommentAuthors = null;
            Stream? source = _pendingSource;
            _pendingSource = null;
            if (source != null && !ReferenceEquals(source, _packageStream)) source.Dispose();
        }
    }
}
