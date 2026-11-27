using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ProjectLegacyMainMasterLayouts(LegacyPptImportParts parts, LegacyPptLayoutCatalog catalog,
            LegacyPptDocument legacy, LegacyPptMaster master, PowerPointPartRef masterPart,
            int masterIndex) {
            LegacyPptSlide[] slides = legacy.Slides.Where(slide => slide.MasterId == master.MasterId)
                .ToArray();
            if (slides.Length == 0) {
                catalog.AddMasterFallback(master.MasterId, new LegacyPptLayoutTarget(masterIndex, 0));
                return;
            }

            PowerPointPartRef firstLayout = LegacyPptImportParts.Related(masterPart, LegacySlideLayoutRelationshipType).First();
            LegacyPptSlide[] representativeSlides = slides
                .GroupBy(LegacyPptLayoutCatalog.CreateKey, StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
            for (int index = 0; index < representativeSlides.Length; index++) {
                LegacyPptSlide representative = representativeSlides[index];
                PowerPointPartRef layoutPart;
                int layoutIndex;
                if (index == 0) {
                    layoutPart = firstLayout;
                    layoutIndex = 0;
                } else {
                    layoutPart = AddLegacyLayoutPart(parts, masterPart);
                    layoutIndex = LegacyPptImportParts.Related(masterPart, LegacySlideLayoutRelationshipType).Count - 1;
                }

                string name = GetLegacySlideLayoutName(master, representative);
                WriteLegacySlideLayout(layoutPart.Root, name, representative,
                    legacy.SlideWidth, legacy.SlideHeight);
                layoutPart.MarkEdited();
                catalog.Add(representative, new LegacyPptLayoutTarget(masterIndex, layoutIndex));
            }
            catalog.AddMasterFallback(master.MasterId, new LegacyPptLayoutTarget(masterIndex, 0));
        }

        /// <summary>
        /// Adds an empty layout part to a master: the master's relationship to it with the next free <c>rIdN</c>, the
        /// layout's relationship back, and its entry in the master's layout list.
        /// </summary>
        private static PowerPointPartRef AddLegacyLayoutPart(LegacyPptImportParts parts, PowerPointPartRef masterPart) {
            OfficeOpcEditablePackage package = parts.Package;
            string relationshipId = LegacyPptImportParts.SmallestFreeRelationshipId(package, masterPart.Name);
            parts.AddXmlPart(masterPart, LegacySlideLayoutRelationshipType, "/ppt/slideLayouts/slideLayout.xml",
                LegacySlideLayoutContentType, LegacyPptXml.P, "sldLayout", _ => { }, out LegacyPptImportPart added,
                relationshipId);
            PowerPointPartRef layoutPart = PowerPointPartRef.Package(parts.Document, added.PartName!);
            package.AddRelationship(layoutPart.Name, LegacySlideMasterRelationshipType,
                masterPart.RelativeTo(layoutPart), id: LegacyPptImportParts.NewSdkRelationshipId());
            OfficeXmlNode slideMaster = masterPart.Root;
            if (!slideMaster.Is(LegacyPptXml.P, "sldMaster")) {
                throw new InvalidDataException("The projected PowerPoint package has no slide master.");
            }
            OfficeXmlNode layoutIds = slideMaster.GetOrAddChild(LegacyPptXml.P, "sldLayoutIdLst", LegacyPptImportXml.SlideMasterOrder);
            // Layout ids are unique across the presentation's masters and layouts, not only within one master.
            uint layoutId = NextLegacyMasterOrLayoutId(parts.PresentationPart, excludedLayoutIds: null);
            layoutIds.AddP("sldLayoutId").Set("id", (long)layoutId).SetRelationship("id", relationshipId);
            masterPart.MarkEdited();
            parts.Document.InvalidatePackageRelationships();
            return layoutPart;
        }

        // Writes a layout root for slides of one binary layout: placeholders, header and footer placeholders and settings.
        private static void WriteLegacySlideLayout(OfficeXmlNode layout, string name,
            LegacyPptSlide source, int slideWidth, int slideHeight) {
            ClearLegacyRoot(layout);
            layout.Set("type", MapLegacyLayoutType(source.Layout, source.LayoutPlaceholderTypes));
            OfficeXmlNode commonSlideData = layout.AddP("cSld").Set("name", name);
            WriteLegacyLayoutPlaceholderTree(commonSlideData.AddP("spTree"), source.Shapes,
                source.HeaderFooter, slideWidth, slideHeight);
            layout.AddP("clrMapOvr").AddA("masterClrMapping");
            ApplyLegacyHeaderFooter(layout, commonSlideData,
                source.HeaderFooter, allowHeader: false);
        }

        // Removes a root's children and attributes other than namespace declarations, as replacing the root does.
        private static void ClearLegacyRoot(OfficeXmlNode root) {
            foreach (OfficeXmlNode child in root.Children().ToList()) child.Remove();
            if (root.Tree is { } tree) {
                for (int index = tree.AttributeCountOf(root.Node) - 1; index >= 0; index--) {
                    tree.GetAttributeAt(root.Node, index, out int namespaceId, out string localName, out _);
                    if (namespaceId == Core.Internal.OpenXml.OfficeXmlNamespaces.Xmlns
                        || namespaceId == Core.Internal.OpenXml.OfficeXmlNamespaces.None && localName == "xmlns") continue;
                    root.RemoveAttribute(namespaceId, localName);
                }
            }
        }

        /// <summary>Writes the group header of a shape tree: its non-visual properties and an empty transform.</summary>
        private static void WriteLegacyShapeTreeHeader(OfficeXmlNode tree) {
            OfficeXmlNode nonVisual = tree.AddP("nvGrpSpPr");
            nonVisual.AddP("cNvPr").Set("id", 1L).Set("name", string.Empty);
            nonVisual.AddP("cNvGrpSpPr");
            nonVisual.AddP("nvPr");
            OfficeXmlNode transform = tree.AddP("grpSpPr").AddA("xfrm");
            transform.AddA("off").Set("x", 0L).Set("y", 0L);
            transform.AddA("ext").Set("cx", 0L).Set("cy", 0L);
            transform.AddA("chOff").Set("x", 0L).Set("y", 0L);
            transform.AddA("chExt").Set("cx", 0L).Set("cy", 0L);
        }

        private static void WriteLegacyLayoutPlaceholderTree(OfficeXmlNode tree,
            IReadOnlyList<LegacyPptShape> shapes,
            LegacyPptHeaderFooterSettings? headerFooter,
            int slideWidth, int slideHeight) {
            WriteLegacyShapeTreeHeader(tree);
            uint shapeId = 2U;
            foreach (LegacyPptShape source in shapes.Where(shape => shape.Placeholder != null)) {
                AppendLegacyLayoutPlaceholderShape(tree, source, shapeId++);
            }
            AppendLegacyHeaderFooterPlaceholders(tree, headerFooter, slideWidth,
                slideHeight, ref shapeId);
        }

        private static void AppendLegacyHeaderFooterPlaceholders(OfficeXmlNode tree,
            LegacyPptHeaderFooterSettings? settings, int slideWidth,
            int slideHeight, ref uint shapeId) {
            if (settings == null) return;
            long width = ToEmus(slideWidth);
            long height = ToEmus(slideHeight);
            long margin = Math.Max(1L, width / 24L);
            long placeholderHeight = Math.Max(1L, height / 16L);
            long top = Math.Max(0L, height - margin - placeholderHeight);
            long sideWidth = Math.Max(1L, width / 5L);
            long centerWidth = Math.Max(1L, width / 3L);
            if (settings.ShowDate || settings.UserDateText.Length > 0) {
                AppendLegacyHeaderFooterPlaceholder(tree, PowerPointPlaceholderType.DateAndTime,
                    settings.UserDateText, margin, top, sideWidth, placeholderHeight,
                    ref shapeId);
            }
            if (settings.ShowFooter || settings.FooterText.Length > 0) {
                AppendLegacyHeaderFooterPlaceholder(tree, PowerPointPlaceholderType.Footer,
                    settings.FooterText, (width - centerWidth) / 2L, top,
                    centerWidth, placeholderHeight, ref shapeId);
            }
            if (settings.ShowSlideNumber) {
                AppendLegacyHeaderFooterPlaceholder(tree, PowerPointPlaceholderType.SlideNumber,
                    string.Empty, width - margin - sideWidth, top, sideWidth,
                    placeholderHeight, ref shapeId);
            }
        }

        private static void AppendLegacyHeaderFooterPlaceholder(OfficeXmlNode tree,
            PowerPointPlaceholderType type, string text, long left, long top, long width,
            long height, ref uint shapeId) {
            string token = type.ToOpenXmlToken();
            bool exists = tree.Children(LegacyPptXml.P, "sp").ToList().Any(shape =>
                shape.Path(LegacyPptXml.P, "nvSpPr", "nvPr", "ph").Attribute("type") == token);
            if (exists) return;
            string name = type == PowerPointPlaceholderType.DateAndTime
                ? "Binary Date Placeholder"
                : type == PowerPointPlaceholderType.Footer
                    ? "Binary Footer Placeholder"
                    : type == PowerPointPlaceholderType.SlideNumber
                        ? "Binary Slide Number Placeholder"
                        : "Binary Header/Footer Placeholder";
            OfficeXmlNode shape = tree.AddP("sp");
            OfficeXmlNode nonVisual = shape.AddP("nvSpPr");
            nonVisual.AddP("cNvPr").Set("id", (long)shapeId++).Set("name", name);
            nonVisual.AddP("cNvSpPr").AddA("spLocks").Set("noGrp", true);
            nonVisual.AddP("nvPr").AddP("ph").Set("type", token);
            AppendLegacyOffsetAndExtents(shape.AddP("spPr").AddA("xfrm"), left, top, width, height);
            OfficeXmlNode textBody = shape.AddP("txBody");
            textBody.AddA("bodyPr");
            textBody.AddA("lstStyle");
            AppendLegacyPlainParagraph(textBody, text);
        }

        private static void AppendLegacyOffsetAndExtents(OfficeXmlNode transform, long left, long top, long width, long height) {
            transform.AddA("off").Set("x", left).Set("y", top);
            transform.AddA("ext").Set("cx", width).Set("cy", height);
        }

        private static void AppendLegacyLayoutPlaceholderShape(OfficeXmlNode tree, LegacyPptShape source,
            uint shapeId) {
            LegacyPptPlaceholder placeholder = source.Placeholder
                ?? throw new InvalidOperationException("A layout placeholder source has no PlaceholderAtom.");
            string name = source.Metadata.Name
                ?? $"Binary {placeholder.Kind} Placeholder {placeholder.Position}";
            OfficeXmlNode shape = tree.AddP("sp");
            OfficeXmlNode nonVisual = shape.AddP("nvSpPr");
            OfficeXmlNode drawing = nonVisual.AddP("cNvPr").Set("id", (long)shapeId).Set("name", name);
            nonVisual.AddP("cNvSpPr").AddA("spLocks").Set("noGrp", true);
            WriteLegacyPlaceholderShape(nonVisual.AddP("nvPr").AddP("ph"), placeholder);
            AppendLegacyOffsetAndExtents(shape.AddP("spPr").AddA("xfrm"),
                ToEmus(source.Bounds.Left), ToEmus(source.Bounds.Top),
                Math.Max(1L, ToEmus(source.Bounds.Width)), Math.Max(1L, ToEmus(source.Bounds.Height)));
            OfficeXmlNode textBody = shape.AddP("txBody");
            textBody.AddA("bodyPr");
            textBody.AddA("lstStyle");
            textBody.AddA("p").AddA("endParaRPr");
            if (source.Metadata.Description != null) {
                drawing.Set("descr", source.Metadata.Description);
            }
        }

        private static void WriteLegacyPlaceholderShape(OfficeXmlNode target,
            LegacyPptPlaceholder source) {
            PowerPointPlaceholderType? type = MapPlaceholder(source.Kind);
            string size = source.Size switch {
                LegacyPptPlaceholderSize.Full => "full",
                LegacyPptPlaceholderSize.Half => "half",
                LegacyPptPlaceholderSize.Quarter => "quarter",
                _ => throw new ArgumentOutOfRangeException(nameof(source))
            };
            uint index = checked((uint)source.Position);
            if (type.HasValue) target.Set("type", type.Value.ToOpenXmlToken());
            if (source.Kind == LegacyPptPlaceholderKind.VerticalTitle
                || source.Kind == LegacyPptPlaceholderKind.VerticalBody
                || source.Kind == LegacyPptPlaceholderKind.VerticalObject) {
                target.Set("orient", "vert");
            }
            target.Set("sz", size).Set("idx", (long)index);
        }

        private static void ApplyLegacyPlaceholder(OfficeXmlNode target,
            LegacyPptShape source) {
            if (source.Placeholder == null || target.IsNone || target.NamespaceId != LegacyPptXml.P) return;
            string? nonVisual = target.LocalName switch {
                "sp" => "nvSpPr",
                "cxnSp" => "nvCxnSpPr",
                "pic" => "nvPicPr",
                "graphicFrame" => "nvGraphicFramePr",
                _ => null
            };
            if (nonVisual == null) return;
            OfficeXmlNode applicationProperties = target.Child(LegacyPptXml.P, nonVisual).Child(LegacyPptXml.P, "nvPr");
            if (applicationProperties.IsNone) return;
            WriteLegacyPlaceholderShape(
                applicationProperties.SetChild(LegacyPptXml.P, "ph", LegacyPptImportXml.ApplicationPropertiesOrder),
                source.Placeholder);
        }

        private static string MapLegacyLayoutType(LegacyPptSlideLayoutType? type,
            IReadOnlyList<LegacyPptPlaceholderKind> placeholderTypes) => type switch {
                LegacyPptSlideLayoutType.TitleSlide => "title",
                LegacyPptSlideLayoutType.TitleBody => MapTitleBodyLayout(placeholderTypes),
                LegacyPptSlideLayoutType.MasterTitle => "title",
                LegacyPptSlideLayoutType.TitleOnly => "titleOnly",
                LegacyPptSlideLayoutType.TwoColumns => "twoColTx",
                LegacyPptSlideLayoutType.TwoRows => "twoObj",
                LegacyPptSlideLayoutType.ColumnTwoRows => "objAndTwoObj",
                LegacyPptSlideLayoutType.TwoRowsColumn => "twoObjAndObj",
                LegacyPptSlideLayoutType.TwoColumnsRow => "twoObjOverTx",
                LegacyPptSlideLayoutType.FourObjects => "fourObj",
                LegacyPptSlideLayoutType.BigObject => "objOnly",
                LegacyPptSlideLayoutType.Blank => "blank",
                LegacyPptSlideLayoutType.VerticalTitleBody => "vertTitleAndTx",
                LegacyPptSlideLayoutType.VerticalTwoRows => "vertTitleAndTxOverChart",
                _ => "cust"
            };

        private static string MapTitleBodyLayout(
            IReadOnlyList<LegacyPptPlaceholderKind> placeholderTypes) {
            LegacyPptPlaceholderKind content = placeholderTypes.Skip(1)
                .FirstOrDefault(value => value != LegacyPptPlaceholderKind.None);
            return content switch {
                LegacyPptPlaceholderKind.Table => "tbl",
                LegacyPptPlaceholderKind.Graph => "chart",
                LegacyPptPlaceholderKind.ClipArt => "txAndClipArt",
                LegacyPptPlaceholderKind.Media => "txAndMedia",
                LegacyPptPlaceholderKind.Object or LegacyPptPlaceholderKind.OrganizationChart
                    or LegacyPptPlaceholderKind.VerticalObject => "txAndObj",
                _ => "tx"
            };
        }

        private static string GetLegacySlideLayoutName(LegacyPptMaster master,
            LegacyPptSlide slide) {
            string type = slide.Layout?.ToString() ?? $"0x{slide.LayoutType:X8}";
            string signature = string.Join("-", slide.LayoutPlaceholderTypes
                .Select(value => ((byte)value).ToString("X2")));
            return $"{GetLegacyMasterName(master)} / {type} / {signature}";
        }

        private sealed class LegacyPptLayoutCatalog {
            private readonly Dictionary<string, LegacyPptLayoutTarget> LoadedSlides =
                new(StringComparer.Ordinal);
            private readonly Dictionary<uint, LegacyPptLayoutTarget> _fallbacks = new();

            internal void Add(LegacyPptSlide slide, LegacyPptLayoutTarget target) =>
                LoadedSlides[CreateKey(slide)] = target;

            internal void AddMasterFallback(uint masterId, LegacyPptLayoutTarget target) =>
                _fallbacks[masterId] = target;

            internal bool TryGet(LegacyPptSlide slide, out LegacyPptLayoutTarget target) =>
                LoadedSlides.TryGetValue(CreateKey(slide), out target)
                || _fallbacks.TryGetValue(slide.MasterId, out target);

            internal static string CreateKey(LegacyPptSlide slide) =>
                $"{slide.MasterId:X8}:{slide.LayoutType:X8}:"
                + string.Join("-", slide.LayoutPlaceholderTypes
                    .Select(value => ((byte)value).ToString("X2")))
                + ":" + (slide.HeaderFooter?.CreateLayoutKey() ?? string.Empty);
        }

        private readonly struct LegacyPptLayoutTarget {
            internal LegacyPptLayoutTarget(int masterIndex, int layoutIndex) {
                MasterIndex = masterIndex;
                LayoutIndex = layoutIndex;
            }

            internal int MasterIndex { get; }
            internal int LayoutIndex { get; }
        }
    }
}
