using System.Collections.ObjectModel;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Captures the synthetic editable layout scope projected from a classic
    /// slide layout signature so supported changes can be materialized without
    /// treating it as an independent binary persist object.
    /// </summary>
    internal sealed class LegacyPptOrdinaryLayoutProjection {
        private readonly IReadOnlyDictionary<uint, string> _shapesById;

        private LegacyPptOrdinaryLayoutProjection(LegacyPptPartView layoutPart,
            IReadOnlyDictionary<uint, string> shapesById) {
            PartUri = layoutPart.Name;
            ShapeTreeFingerprint = CreateShapeTreeFingerprint(layoutPart);
            TypeFingerprint = CreateTypeFingerprint(layoutPart);
            _shapesById = new ReadOnlyDictionary<uint, string>(
                shapesById.ToDictionary(pair => pair.Key,
                    pair => pair.Value));
        }

        internal string PartUri { get; }

        internal string ShapeTreeFingerprint { get; }

        internal string TypeFingerprint { get; }

        internal bool ShapeTreeMatches(LegacyPptPartView layoutPart) =>
            string.Equals(ShapeTreeFingerprint,
                CreateShapeTreeFingerprint(layoutPart),
                StringComparison.Ordinal);

        internal bool TypeMatches(LegacyPptPartView layoutPart) =>
            string.Equals(TypeFingerprint,
                CreateTypeFingerprint(layoutPart),
                StringComparison.Ordinal);

        internal bool TryGetAddedShapeIds(LegacyPptPartView layoutPart,
            out IReadOnlyList<uint> addedShapeIds) {
            addedShapeIds = Array.Empty<uint>();
            if (!TryCreateShapeSnapshot(layoutPart,
                    out IReadOnlyDictionary<uint, string> current)) {
                return false;
            }
            foreach (KeyValuePair<uint, string> expected in _shapesById) {
                if (!current.TryGetValue(expected.Key,
                        out string? currentShape)
                    || !string.Equals(expected.Value, currentShape,
                        StringComparison.Ordinal)) {
                    return false;
                }
            }
            addedShapeIds = current.Keys.Where(id =>
                    !_shapesById.ContainsKey(id))
                .ToArray();
            return true;
        }

        internal static IReadOnlyList<LegacyPptOrdinaryLayoutProjection>
            Create(LegacyPptPackageView view,
                IReadOnlyList<LegacyPptMasterProjection> titleMasters) {
            ArgumentNullException.ThrowIfNull(view);
            var titlePartUris = new HashSet<string>(titleMasters.Select(
                master => master.MasterPartUri), StringComparer.Ordinal);
            var result = new List<LegacyPptOrdinaryLayoutProjection>();
            foreach (LegacyPptPartView layout in LegacyPptPackageFingerprint.SlideMasterParts(view).SelectMany(
                         LegacyPptPackageFingerprint.SlideLayoutParts)) {
                if (titlePartUris.Contains(layout.Name)
                    || !TryCreateShapeSnapshot(layout,
                        out IReadOnlyDictionary<uint, string> shapes)) {
                    continue;
                }
                result.Add(new LegacyPptOrdinaryLayoutProjection(layout,
                    shapes));
            }
            return result;
        }

        private static OfficeXmlNode ShapeTree(LegacyPptPartView layoutPart) =>
            layoutPart.Root.Path(LegacyPptXml.P, "cSld", "spTree");

        private static string CreateShapeTreeFingerprint(
            LegacyPptPartView layoutPart) => LegacyPptMasterProjection.Markup(ShapeTree(layoutPart));

        private static string CreateTypeFingerprint(
            LegacyPptPartView layoutPart) => layoutPart.Root.Attribute("type") ?? string.Empty;

        private static bool TryCreateShapeSnapshot(
            LegacyPptPartView layoutPart,
            out IReadOnlyDictionary<uint, string> snapshot) {
            var result = new Dictionary<uint, string>();
            snapshot = result;
            foreach (OfficeXmlNode element in ShapeTree(layoutPart).Children()) {
                if (element.Is(LegacyPptXml.P, "nvGrpSpPr")
                    || element.Is(LegacyPptXml.P, "grpSpPr")) continue;
                uint? id = GetShapeId(element);
                if (!id.HasValue || result.ContainsKey(id.Value)) {
                    return false;
                }
                OfficeXmlNode normalized = element.Clone();
                if (normalized.Is(LegacyPptXml.P, "sp")
                    && IsHeaderFooterPlaceholder(normalized)) {
                    OfficeXmlNode text = normalized.Child(LegacyPptXml.P, "txBody");
                    if (!text.IsNone) text.RemoveChildren(LegacyPptXml.A, "p");
                }
                result.Add(id.Value, LegacyPptXml.Digest(normalized));
            }
            return true;
        }

        /// <summary>The <c>p:cNvPr</c> id of a shape-tree element, or null for an element that is not a shape.</summary>
        internal static uint? GetShapeId(OfficeXmlNode element) {
            if (element.NamespaceId != LegacyPptXml.P) return null;
            string? container = element.LocalName switch {
                "sp" => "nvSpPr",
                "pic" => "nvPicPr",
                "cxnSp" => "nvCxnSpPr",
                "grpSp" => "nvGrpSpPr",
                "graphicFrame" => "nvGraphicFramePr",
                _ => null
            };
            return container == null
                ? null
                : element.Path(LegacyPptXml.P, container, "cNvPr").UInt("id");
        }

        private static bool IsHeaderFooterPlaceholder(OfficeXmlNode shape) =>
            shape.Path(LegacyPptXml.P, "nvSpPr", "nvPr", "ph").Attribute("type")
                is "dt" or "ftr" or "sldNum";
    }
}
