using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The chart part behind a chart frame and its elements, read through <see cref="PowerPointPartRef"/> and
    /// <see cref="OfficeXmlNode"/>, so a chart loaded from a file is read and edited without opening the SDK.
    /// </summary>
    public partial class PowerPointChart {
        private const int CNs = OfficeXmlNamespaces.Chart;
        private const int ANs = OfficeXmlNamespaces.DrawingML;

        /// <summary>The chart part the frame's <c>c:chart</c> reference points to, from the part that holds the frame.</summary>
        internal PowerPointPartRef ChartPartRef {
            get {
                OfficeXmlNode reference = XmlElement.Child(ANs, "graphic").Child(ANs, "graphicData").Child(CNs, "chart");
                string? relationshipId = reference.Attribute(OfficeXmlNamespaces.Relationships, "id");
                if (relationshipId == null) {
                    throw new InvalidOperationException("Chart reference not found for the shape.");
                }
                if (!OwnerPartRef.TryGetRelationship(relationshipId, out PowerPointRelationship relationship) || relationship.TargetPart.IsNone) {
                    throw new ArgumentOutOfRangeException(nameof(relationshipId), "The chart relationship '" + relationshipId + "' was not found.");
                }
                return relationship.TargetPart;
            }
        }

        /// <summary>The chart part's <c>c:chartSpace</c>, or none when the part has no content.</summary>
        internal OfficeXmlNode ChartSpaceNode => ChartPartRef.Root;

        /// <summary>The chart space's <c>c:chart</c>.</summary>
        /// <exception cref="InvalidOperationException">The chart part has no chart element.</exception>
        internal OfficeXmlNode ChartNode {
            get {
                OfficeXmlNode chart = ChartSpaceNode.Child(CNs, "chart");
                if (chart.IsNone) {
                    throw new InvalidOperationException("Chart element not found in chart part.");
                }
                return chart;
            }
        }

        /// <summary>The chart's <c>c:plotArea</c>, or none.</summary>
        internal OfficeXmlNode PlotAreaNode => ChartNode.Child(CNs, "plotArea");

        /// <summary>Records an edit of the chart space, so the document writes it.</summary>
        internal void MarkChartEdited() => ChartPartRef.SaveRootEdited();

        /// <summary>The chart part as the shared chart reader and writer see it.</summary>
        internal IOfficeChartPart SharedChartPart => new PowerPointChartPart(ChartPartRef, OwnerPartRef);
    }

    /// <summary>A chart part of the presentation for the shared chart reader and writer, wherever the part is held.</summary>
    internal sealed class PowerPointChartPart : IOfficeChartPart {
        internal const string ChartStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
        internal const string PackageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";

        private readonly PowerPointPartRef _part;
        private readonly PowerPointPartRef _owner;

        /// <param name="part">The chart part.</param>
        /// <param name="owner">The slide, layout, master or notes part that holds the chart.</param>
        internal PowerPointChartPart(PowerPointPartRef part, PowerPointPartRef owner) {
            _part = part;
            _owner = owner;
        }

        public OfficeXmlNode ChartSpace => _part.Root;

        public IReadOnlyList<OfficeXmlNode> ChartStyles(int limit) {
            var styles = new List<OfficeXmlNode>();
            if (limit <= 0) return styles;
            foreach (PowerPointPartRef style in _part.RelatedParts(ChartStyleRelationshipType)) {
                styles.Add(style.Root);
                if (styles.Count >= limit) break;
            }
            return styles;
        }

        // Slide override, then layout override, then master map, as the SDK chart part reads it from its parent part.
        public OfficeXmlNode HostColorMap {
            get {
                OfficeXmlNode root = _owner.Root;
                switch (_owner.Kind) {
                    case PowerPointPartKind.Slide: {
                        OfficeXmlNode map = OverrideMap(root);
                        if (map.Exists) return map;
                        PowerPointPartRef layout = _owner.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
                        return LayoutMap(layout);
                    }
                    case PowerPointPartKind.SlideLayout:
                        return LayoutMap(_owner);
                    case PowerPointPartKind.SlideMaster:
                        return root.Child(OfficeXmlNamespaces.PresentationML, "clrMap");
                    case PowerPointPartKind.NotesSlide: {
                        OfficeXmlNode map = OverrideMap(root);
                        return map.Exists ? map
                            : _owner.RelatedPart(NotesMasterRelationshipType).Root.Child(OfficeXmlNamespaces.PresentationML, "clrMap");
                    }
                    default:
                        // A notes or handout master keeps its own map.
                        return root.Child(OfficeXmlNamespaces.PresentationML, "clrMap");
                }
            }
        }

        private const string NotesMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";

        private static OfficeXmlNode OverrideMap(OfficeXmlNode root) =>
            root.Child(OfficeXmlNamespaces.PresentationML, "clrMapOvr").Child(OfficeXmlNamespaces.DrawingML, "overrideClrMapping");

        private static OfficeXmlNode LayoutMap(PowerPointPartRef layout) {
            if (layout.IsNone) return default;
            OfficeXmlNode map = OverrideMap(layout.Root);
            return map.Exists ? map
                : layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType).Root.Child(OfficeXmlNamespaces.PresentationML, "clrMap");
        }

        public OfficeXmlNode HostFontScheme {
            get {
                switch (_owner.Kind) {
                    case PowerPointPartKind.Slide:
                    case PowerPointPartKind.SlideLayout:
                    case PowerPointPartKind.SlideMaster:
                        return PowerPointThemeScope.For(_owner).FontScheme;
                    case PowerPointPartKind.NotesSlide: {
                        OfficeXmlNode scheme = _owner.RelatedPart(PowerPointPartRef.ThemeOverrideRelationshipType).Root.Child(OfficeXmlNamespaces.DrawingML, "fontScheme");
                        return scheme.Exists ? scheme : ThemeFontScheme(_owner.RelatedPart(NotesMasterRelationshipType));
                    }
                    default:
                        return ThemeFontScheme(_owner);
                }
            }
        }

        private static OfficeXmlNode ThemeFontScheme(PowerPointPartRef part) =>
            part.IsNone ? default : part.RelatedPart(PowerPointPartRef.ThemeRelationshipType).Root
                .Child(OfficeXmlNamespaces.DrawingML, "themeElements").Child(OfficeXmlNamespaces.DrawingML, "fontScheme");

        public IReadOnlyList<IOfficeChartEmbeddedPackage> EmbeddedPackages(int limit) {
            var packages = new List<IOfficeChartEmbeddedPackage>();
            if (limit <= 0) return packages;
            foreach (KeyValuePair<string, PowerPointPartRef> related in _part.RelatedPartsById(PackageRelationshipType)) {
                packages.Add(new EmbeddedPackage(related.Key, related.Value));
                if (packages.Count >= limit) break;
            }
            return packages;
        }

        public void MarkEdited() => _part.SaveRootEdited();

        private sealed class EmbeddedPackage : IOfficeChartEmbeddedPackage {
            private readonly PowerPointPartRef _package;

            internal EmbeddedPackage(string relationshipId, PowerPointPartRef package) {
                RelationshipId = relationshipId;
                _package = package;
            }

            public string RelationshipId { get; }

            public string ContentType => _package.ContentType ?? string.Empty;

            public Stream OpenRead() => new MemoryStream(_package.ReadBytes(), writable: false);

            public void Replace(byte[] data) => _package.ReplaceBytes(data);
        }
    }
}
