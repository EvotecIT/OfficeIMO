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
        internal IOfficeChartPart SharedChartPart {
            get {
                var part = new PowerPointChartPart(ChartPartRef, OwnerPartRef);
                // A chart read from the package tells the shared reader which part its tree is, so chart text and automatic
                // colours find the host color map and chart styles as they do for an SDK part.
                if (part.ChartSpace.Tree is { } tree) OfficeOpenXmlChartTextReader.AttachChartPart(tree, part);
                return part;
            }
        }

        /// <summary>The theme color scheme in effect for the part that holds the chart, or none.</summary>
        internal OfficeXmlNode OwnerColorScheme => PowerPointChartPart.HostScheme(OwnerPartRef, "clrScheme");
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

        public IReadOnlyList<Func<Stream>> ChartStyles(int limit) {
            var styles = new List<Func<Stream>>();
            if (limit <= 0) return styles;
            foreach (PowerPointPartRef style in _part.RelatedParts(ChartStyleRelationshipType)) {
                // The stored bytes, without parsing the part first, so the reader's size limit applies as it reads.
                styles.Add(() => new MemoryStream(style.ReadBytes(), writable: false));
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

        public OfficeXmlNode HostFontScheme => HostScheme(_owner, "fontScheme");

        /// <summary>
        /// The theme scheme (<c>clrScheme</c>, <c>fontScheme</c>) in effect for <paramref name="owner"/>: its theme override,
        /// then its layout's, then its master's theme; a notes slide's override, then its notes master's theme; a master's
        /// own theme.
        /// </summary>
        internal static OfficeXmlNode HostScheme(PowerPointPartRef owner, string schemeName) {
            // Each part is read only when the ones before it have no scheme, as the SDK chain reads them.
            OfficeXmlNode scheme;
            switch (owner.Kind) {
                case PowerPointPartKind.Slide: {
                    if ((scheme = OverrideScheme(owner, schemeName)).Exists) return scheme;
                    PowerPointPartRef layout = owner.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
                    if (layout.IsNone) return default;
                    if ((scheme = OverrideScheme(layout, schemeName)).Exists) return scheme;
                    return ThemeScheme(layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType), schemeName);
                }
                case PowerPointPartKind.SlideLayout:
                    if ((scheme = OverrideScheme(owner, schemeName)).Exists) return scheme;
                    return ThemeScheme(owner.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType), schemeName);
                case PowerPointPartKind.NotesSlide:
                    if ((scheme = OverrideScheme(owner, schemeName)).Exists) return scheme;
                    return ThemeScheme(owner.RelatedPart(NotesMasterRelationshipType), schemeName);
                default:
                    return ThemeScheme(owner, schemeName);
            }
        }

        private static OfficeXmlNode OverrideScheme(PowerPointPartRef part, string schemeName) =>
            RootOf(part.RelatedPart(PowerPointPartRef.ThemeOverrideRelationshipType)).Child(OfficeXmlNamespaces.DrawingML, schemeName);

        private static OfficeXmlNode ThemeScheme(PowerPointPartRef part, string schemeName) =>
            part.IsNone ? default : RootOf(part.RelatedPart(PowerPointPartRef.ThemeRelationshipType))
                .Child(OfficeXmlNamespaces.DrawingML, "themeElements").Child(OfficeXmlNamespaces.DrawingML, schemeName);

        // A part's root; a package part that exists but can't be parsed fails, as reading it through the SDK does.
        private static OfficeXmlNode RootOf(PowerPointPartRef part) {
            if (part.IsNone) return default;
            OfficeXmlNode root = part.Root;
            if (root.IsNone && !part.Current.IsSdk) throw new InvalidDataException("The part '" + part.Name + "' could not be read.");
            return root;
        }

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
