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
            // Notes are rewritten whole by the model, and a part size limit is the SDK's to enforce.
            if (_disposed || !_sdkDeferred || _sdkDocument != null || _accessMode != DocumentAccessMode.ReadWrite
                || options?.IncludeNotes == true || _deferredOpenSettings?.MaxCharactersInPart > 0
                || _packageStream is not { CanRead: true, CanSeek: true } stream) {
                return false;
            }

            OfficeOpcEditablePackage package;
            if (_pendingPackage != null) {
                package = _pendingPackage;
            } else {
                stream.Position = 0;
                package = _pendingPackage = OfficeOpcEditablePackage.Open(stream, leaveOpen: true);
            }
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

        // The slide parts of the first slide list, in order; null when one is missing, external or listed twice.
        private static List<string>? ReadSlideParts(OfficeOpcEditablePackage package) {
            string? main = null;
            OfficeXmlTree packageRelationships = package.GetRelationships(null);
            for (int node = packageRelationships.FindChild(packageRelationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship");
                 node != OfficeXmlTree.NoNode;
                 node = packageRelationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string? type = packageRelationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                if (type != null && OfficeXmlNamespaces.NormalizeRelationshipType(type) == OfficeOpcPackage.OfficeDocumentRelationshipType) {
                    main = OfficeOpcPartNames.ResolveRelationshipTarget("/_rels/.rels", packageRelationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                    break;
                }
            }
            if (main == null || !package.PartExists(main)) return null;
            OfficeXmlTree presentation = package.GetXml(main);
            int list = presentation.FindChild(presentation.Root, OfficeXmlNamespaces.PresentationML, "sldIdLst");
            var slides = new List<string>();
            if (list == OfficeXmlTree.NoNode) return slides;
            var seen = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
            for (int entry = presentation.FindChild(list, OfficeXmlNamespaces.PresentationML, "sldId"); entry != OfficeXmlTree.NoNode;
                 entry = presentation.FindNextSibling(entry, OfficeXmlNamespaces.PresentationML, "sldId")) {
                string? id = presentation.GetAttribute(entry, OfficeXmlNamespaces.Relationships, "id");
                string? slide = id == null ? null : package.GetRelationshipTarget(main, id);
                if (slide == null || !package.PartExists(slide) || !seen.Add(OfficeOpcPackage.GetPartIdentity(slide))) return null;
                slides.Add(slide);
            }
            return slides;
        }

        /// <summary>Writes pending slide edits into the package bytes, so whatever reads them next sees the edits.</summary>
        private void CommitPendingEdits() {
            OfficeOpcEditablePackage? pending = _pendingPackage;
            if (pending == null) return;
            _pendingPackage = null;
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
            _pendingEdits = false;
        }
    }
}
