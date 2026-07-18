using System.Collections.ObjectModel;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordEndDocument = 0x03EA;
        private const ushort RecordSlideShowDocInfoAtom = 0x0401;
        private const ushort RecordSummary = 0x0402;
        private const ushort RecordNamedShows = 0x0410;
        private const ushort RecordNamedShow = 0x0411;
        private const ushort RecordNamedShowSlides = 0x0412;

        internal static bool TryReadCustomShows(PowerPointPresentation presentation,
            out LegacyPptWriterCustomShowCatalog catalog, out string? reason) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            catalog = new LegacyPptWriterCustomShowCatalog();
            reason = null;
            PresentationPart? presentationPart = presentation.OpenXmlDocument
                .PresentationPart;
            P.CustomShowList? list = presentationPart?.Presentation?.CustomShowList;
            if (list == null) return true;
            if (list.HasAttributes || list.ChildElements.Any(child =>
                    child is not P.CustomShow)) {
                reason = "The custom-show list contains attributes or extension data that have no binary representation.";
                return false;
            }

            var ids = new HashSet<uint>();
            var names = new HashSet<string>(StringComparer.Ordinal);
            var slideUris = new HashSet<string>(presentation.Slides.Select(slide =>
                slide.SlidePart.Uri.ToString()), StringComparer.Ordinal);
            foreach (P.CustomShow show in list.Elements<P.CustomShow>()) {
                string? name = show.Name?.Value;
                uint? id = show.Id?.Value;
                if (name == null || name.Length == 0 || name.IndexOf('\0') >= 0
                    || !id.HasValue || !ids.Add(id.Value)
                    || !names.Add(name)) {
                    reason = "Custom shows require unique numeric ids, unique nonempty names, and valid Unicode text.";
                    return false;
                }
                if (show.GetAttributes().Any(attribute =>
                        !string.Equals(attribute.LocalName, "name",
                            StringComparison.Ordinal)
                        && !string.Equals(attribute.LocalName, "id",
                            StringComparison.Ordinal))
                    || show.ChildElements.Any(child => child is not P.SlideList)
                    || show.Elements<P.SlideList>().Skip(1).Any()) {
                    reason = $"Custom show '{name}' contains extension data that have no binary representation.";
                    return false;
                }
                P.SlideList? showSlides = show.SlideList;
                if (showSlides != null && (showSlides.HasAttributes
                        || showSlides.ChildElements.Any(child =>
                            child is not P.SlideListEntry))) {
                    reason = $"Custom show '{name}' contains unsupported slide-list data.";
                    return false;
                }

                var showSlideUris = new List<string>();
                foreach (P.SlideListEntry entry in showSlides?
                             .Elements<P.SlideListEntry>()
                         ?? Enumerable.Empty<P.SlideListEntry>()) {
                    string? relationshipId = entry.Id?.Value;
                    if (entry.GetAttributes().Any(attribute =>
                            !string.Equals(attribute.LocalName, "id",
                                StringComparison.Ordinal))
                        || string.IsNullOrEmpty(relationshipId)
                        || !presentationPart!.TryGetPartById(relationshipId!,
                            out OpenXmlPart? part)
                        || part is not SlidePart slidePart
                        || !slideUris.Contains(slidePart.Uri.ToString())) {
                        reason = $"Custom show '{name}' references a missing or unsupported slide relationship.";
                        return false;
                    }
                    showSlideUris.Add(slidePart.Uri.ToString());
                }
                catalog.Add(new LegacyPptWriterCustomShow(name, showSlideUris));
            }
            return true;
        }

        internal static bool TryBuildNamedShowsRecord(
            LegacyPptWriterCustomShowCatalog catalog,
            Func<string, uint?> resolveSlideId, out byte[] bytes) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            if (resolveSlideId == null) throw new ArgumentNullException(nameof(resolveSlideId));
            if (catalog.Shows.Count == 0) {
                bytes = Array.Empty<byte>();
                return true;
            }
            var shows = new List<byte[]>(catalog.Shows.Count);
            foreach (LegacyPptWriterCustomShow show in catalog.Shows) {
                var children = new List<byte[]> {
                    BuildRecord(version: 0, instance: 0, RecordCString,
                        Encoding.Unicode.GetBytes(show.Name))
                };
                if (show.SlidePartUris.Count > 0) {
                    var payload = new byte[checked(show.SlidePartUris.Count * 4)];
                    for (int index = 0; index < show.SlidePartUris.Count; index++) {
                        uint? slideId = resolveSlideId(show.SlidePartUris[index]);
                        if (!slideId.HasValue) {
                            bytes = Array.Empty<byte>();
                            return false;
                        }
                        WriteUInt32(payload, index * 4, slideId.Value);
                    }
                    children.Add(BuildRecord(version: 0, instance: 0,
                        RecordNamedShowSlides, payload));
                }
                shows.Add(BuildContainer(RecordNamedShow, instance: 0,
                    children));
            }
            bytes = BuildContainer(RecordNamedShows, instance: 0, shows);
            return true;
        }

        internal static bool TryRewriteDocumentNamedShows(LegacyPptRecord document,
            byte[] namedShowsRecord, out byte[] bytes) {
            bytes = document.CopyRecordBytes();
            if (document.Type != RecordDocument || document.Version != 0x0F) {
                return false;
            }
            LegacyPptRecord[] existing = document.Children.Where(child =>
                child.Type == RecordNamedShows).ToArray();
            if (existing.Length > 1) return false;

            var children = new List<byte[]>(document.Children.Count
                + (namedShowsRecord.Length == 0 ? 0 : 1));
            bool inserted = false;
            for (int index = 0; index < document.Children.Count; index++) {
                LegacyPptRecord child = document.Children[index];
                if (child.Type == RecordNamedShows) {
                    if (namedShowsRecord.Length > 0) children.Add(namedShowsRecord);
                    inserted = true;
                    continue;
                }
                if (!inserted && existing.Length == 0 && namedShowsRecord.Length > 0
                    && IsAfterNamedShowsBoundary(child)) {
                    children.Add(namedShowsRecord);
                    inserted = true;
                }
                children.Add(child.CopyRecordBytes());
                if (!inserted && existing.Length == 0 && namedShowsRecord.Length > 0
                    && child.Type == RecordSlideShowDocInfoAtom) {
                    children.Add(namedShowsRecord);
                    inserted = true;
                }
            }
            if (!inserted && namedShowsRecord.Length > 0) {
                children.Add(namedShowsRecord);
            }
            bytes = BuildRecord(document.Version, document.Instance,
                document.Type, Concat(children));
            return true;
        }

        private static bool IsAfterNamedShowsBoundary(LegacyPptRecord child) =>
            child.Type == RecordSummary || child.Type == RecordEndDocument;

        internal sealed class LegacyPptWriterCustomShowCatalog {
            private readonly List<LegacyPptWriterCustomShow> _shows = new();

            internal IReadOnlyList<LegacyPptWriterCustomShow> Shows =>
                new ReadOnlyCollection<LegacyPptWriterCustomShow>(_shows);

            internal void Add(LegacyPptWriterCustomShow show) {
                _shows.Add(show);
            }
        }

        internal sealed class LegacyPptWriterCustomShow {
            internal LegacyPptWriterCustomShow(string name,
                IReadOnlyList<string> slidePartUris) {
                Name = name;
                SlidePartUris = new ReadOnlyCollection<string>(
                    slidePartUris.ToArray());
            }

            internal string Name { get; }
            internal IReadOnlyList<string> SlidePartUris { get; }
        }
    }
}
