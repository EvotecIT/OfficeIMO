using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>An Open XML SDK chart part for the shared chart reader and writer.</summary>
    internal sealed class OfficeSdkChartPart : IOfficeChartPart {
        internal OfficeSdkChartPart(ChartPart part) => Part = part ?? throw new ArgumentNullException(nameof(part));

        internal ChartPart Part { get; }

        public OfficeXmlNode ChartSpace => OfficeXmlNode.FromSdk(Part.ChartSpace);

        public IReadOnlyList<OfficeXmlNode> ChartStyles(int limit) =>
            Part.GetPartsOfType<ChartStylePart>().Take(limit).Select(style => OfficeXmlNode.FromSdk(style.RootElement)).ToList();

        public OfficeXmlNode HostColorMap {
            get {
                foreach (OpenXmlPart owner in Part.GetParentParts()) {
                    if (owner is SlidePart slide)
                        return OfficeXmlNode.FromSdk(slide.Slide?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                            ?? slide.SlideLayoutPart?.SlideLayout?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                            ?? (OpenXmlElement?)slide.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap);
                    if (owner is SlideLayoutPart layout)
                        return OfficeXmlNode.FromSdk(layout.SlideLayout?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                            ?? (OpenXmlElement?)layout.SlideMasterPart?.SlideMaster?.ColorMap);
                    if (owner is SlideMasterPart master)
                        return OfficeXmlNode.FromSdk(master.SlideMaster?.ColorMap);
                    if (owner is NotesSlidePart notes)
                        return OfficeXmlNode.FromSdk(notes.NotesSlide?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                            ?? (OpenXmlElement?)notes.NotesMasterPart?.NotesMaster?.ColorMap);
                    if (owner is NotesMasterPart notesMaster)
                        return OfficeXmlNode.FromSdk(notesMaster.NotesMaster?.ColorMap);
                    if (owner is HandoutMasterPart handoutMaster)
                        return OfficeXmlNode.FromSdk(handoutMaster.HandoutMaster?.ColorMap);
                    if (owner is MainDocumentPart document)
                        return OfficeXmlNode.FromSdk(document.DocumentSettingsPart?.Settings?
                            .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.ColorSchemeMapping>());
                    foreach (OpenXmlPart parent in owner.GetParentParts())
                        if (parent is MainDocumentPart parentDocument)
                            return OfficeXmlNode.FromSdk(parentDocument.DocumentSettingsPart?.Settings?
                                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.ColorSchemeMapping>());
                }
                return default;
            }
        }

        public OfficeXmlNode HostFontScheme {
            get {
                var visited = new HashSet<OpenXmlPart>();
                var queue = new Queue<OpenXmlPart>(Part.GetParentParts());
                while (queue.Count > 0 && visited.Count < 16) {
                    OpenXmlPart owner = queue.Dequeue();
                    if (!visited.Add(owner)) continue;
                    A.FontScheme? scheme = owner switch {
                        SlidePart slide => slide.ThemeOverridePart?.ThemeOverride?.FontScheme ??
                            slide.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.FontScheme ??
                            slide.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme,
                        SlideLayoutPart layout => layout.ThemeOverridePart?.ThemeOverride?.FontScheme ??
                            layout.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme,
                        SlideMasterPart master => master.ThemePart?.Theme?.ThemeElements?.FontScheme,
                        NotesSlidePart notes => notes.ThemeOverridePart?.ThemeOverride?.FontScheme ??
                            notes.NotesMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme,
                        NotesMasterPart notesMaster => notesMaster.ThemePart?.Theme?.ThemeElements?.FontScheme,
                        HandoutMasterPart handout => handout.ThemePart?.Theme?.ThemeElements?.FontScheme,
                        MainDocumentPart document => document.ThemePart?.Theme?.ThemeElements?.FontScheme,
                        WorkbookPart workbook => workbook.ThemePart?.Theme?.ThemeElements?.FontScheme,
                        _ => null
                    };
                    if (scheme != null) return OfficeXmlNode.FromSdk(scheme);
                    foreach (OpenXmlPart parent in owner.GetParentParts()) queue.Enqueue(parent);
                }
                return default;
            }
        }

        public IReadOnlyList<IOfficeChartEmbeddedPackage> EmbeddedPackages(int limit) {
            var packages = new List<IOfficeChartEmbeddedPackage>();
            foreach (IdPartPair pair in Part.Parts) {
                if (packages.Count >= limit) break;
                if (pair.OpenXmlPart is EmbeddedPackagePart package) packages.Add(new EmbeddedPackage(pair.RelationshipId, package));
            }
            return packages;
        }

        public void MarkEdited() => Part.ChartSpace?.Save();

        private sealed class EmbeddedPackage : IOfficeChartEmbeddedPackage {
            private readonly EmbeddedPackagePart _part;

            internal EmbeddedPackage(string relationshipId, EmbeddedPackagePart part) {
                RelationshipId = relationshipId;
                _part = part;
            }

            public string RelationshipId { get; }

            public string ContentType => _part.ContentType;

            public Stream OpenRead() => _part.GetStream(FileMode.Open, FileAccess.Read);

            public void Replace(byte[] data) {
                using var stream = new MemoryStream(data, writable: false);
                _part.FeedData(stream);
            }
        }
    }
}
