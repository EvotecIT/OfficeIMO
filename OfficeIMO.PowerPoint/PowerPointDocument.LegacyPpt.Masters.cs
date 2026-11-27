using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const int LegacyP = LegacyPptXml.P;

        private static LegacyPptLayoutCatalog ProjectLegacyMasters(
            LegacyPptImportParts parts, LegacyPptDocument legacy,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            LegacyPptMaster[] mainMasters = legacy.Masters.Where(master => master.IsMainMaster).ToArray();
            var catalog = new LegacyPptLayoutCatalog();
            if (mainMasters.Length == 0) {
                return catalog;
            }

            PowerPointPartRef presentationPart = parts.PresentationPart;
            PowerPointPartRef firstMasterPart = LegacyPptImportParts.Related(presentationPart, LegacySlideMasterRelationshipType).First();
            ResetLegacyMasterScaffold(parts, firstMasterPart, "Binary Main Master 1");
            var masterParts = new List<PowerPointPartRef> { firstMasterPart };
            for (int index = 1; index < mainMasters.Length; index++) {
                masterParts.Add(CloneLegacyMasterScaffold(parts, presentationPart, firstMasterPart));
            }

            var masterTargets = new Dictionary<uint, LegacyPptLayoutTarget>();
            for (int index = 0; index < mainMasters.Length; index++) {
                LegacyPptMaster mainMaster = mainMasters[index];
                PowerPointPartRef masterPart = masterParts[index];
                ApplyLegacyMaster(parts, masterPart, mainMaster,
                    GetEffectiveLegacySlideHeaderFooter(legacy, mainMaster),
                    soundContext, deferredInteractions);
                var target = new LegacyPptLayoutTarget(index, 0);
                masterTargets.Add(mainMaster.MasterId, target);
                ProjectLegacyMainMasterLayouts(parts, catalog, legacy, mainMaster,
                    masterPart, index);
            }

            foreach (LegacyPptMaster titleMaster in legacy.Masters.Where(master => !master.IsMainMaster)) {
                LegacyPptLayoutTarget parentTarget;
                if (!masterTargets.TryGetValue(titleMaster.ParentMasterId, out parentTarget)) {
                    parentTarget = new LegacyPptLayoutTarget(0, 0);
                }
                PowerPointPartRef parentPart = masterParts[parentTarget.MasterIndex];
                int layoutIndex = AddLegacyTitleMasterLayout(parts, parentPart, titleMaster,
                    soundContext, deferredInteractions);
                var titleTarget = new LegacyPptLayoutTarget(parentTarget.MasterIndex, layoutIndex);
                masterTargets[titleMaster.MasterId] = titleTarget;
                catalog.AddMasterFallback(titleMaster.MasterId, titleTarget);
            }

            return catalog;
        }

        // Keeps the scaffold master's first layout as an empty blank layout and removes the others.
        private static void ResetLegacyMasterScaffold(LegacyPptImportParts parts, PowerPointPartRef masterPart, string name) {
            OfficeOpcEditablePackage package = parts.Package;
            List<PowerPointPartRef> layouts = LegacyPptImportParts.Related(masterPart, LegacySlideLayoutRelationshipType);
            if (layouts.Count == 0) {
                throw new InvalidDataException("The native PowerPoint scaffold has no slide layout.");
            }
            PowerPointPartRef blankLayout = layouts[0];
            foreach (PowerPointPartRef extraLayout in layouts.Skip(1)) {
                RemoveLegacyRelationshipsTo(package, masterPart.Name, extraLayout.Name);
                package.DeletePart(OfficeOpcPartNames.GetRelationshipsPartName(extraLayout.Name));
                package.DeletePart(extraLayout.Name);
            }
            parts.Document.InvalidatePackageRelationships();

            WriteLegacyLayout(blankLayout, "Binary Main Master", "blank",
                Array.Empty<LegacyPptShape>());
            OfficeXmlNode slideMaster = masterPart.Root;
            if (!slideMaster.Is(LegacyP, "sldMaster")) {
                throw new InvalidDataException("The native PowerPoint scaffold has no slide master.");
            }
            OfficeXmlNode commonSlideData = slideMaster.SetChild(LegacyP, "cSld", LegacyPptImportXml.SlideMasterOrder).Set("name", name);
            WriteLegacyShapeTree(commonSlideData.AddP("spTree"));
            string? relationshipId = parts.RelationshipIdTo(masterPart, blankLayout);
            slideMaster.SetChild(LegacyP, "sldLayoutIdLst", LegacyPptImportXml.SlideMasterOrder)
                .AddP("sldLayoutId").Set("id", 2147483649L).SetRelationship("id", relationshipId);
            blankLayout.MarkEdited();
            masterPart.MarkEdited();
        }

        private static void RemoveLegacyRelationshipsTo(OfficeOpcEditablePackage package, string owner, string target) {
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(owner);
            if (!package.PartExists(relationshipsPart)) return;
            OfficeXmlTree tree = package.GetXml(relationshipsPart);
            var remove = new List<int>();
            for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                string? resolved = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                if (resolved != null && string.Equals(resolved, target, StringComparison.OrdinalIgnoreCase)) remove.Add(node);
            }
            foreach (int node in remove) tree.Remove(node);
        }

        /// <summary>
        /// Adds a copy of the scaffold master with its layouts and theme, listed in the presentation, as the engine clones a
        /// master: the layouts and theme keep their relationship ids and the copied layouts get new layout ids.
        /// </summary>
        private static PowerPointPartRef CloneLegacyMasterScaffold(LegacyPptImportParts parts,
            PowerPointPartRef presentationPart, PowerPointPartRef sourceMaster) {
            OfficeOpcEditablePackage package = parts.Package;
            PowerPointDocument document = parts.Document;
            string masterRelationshipId = LegacyPptImportParts.SmallestFreeRelationshipId(package, presentationPart.Name);
            parts.AddXmlPart(presentationPart, LegacySlideMasterRelationshipType, "/ppt/slideMasters/slideMaster.xml",
                LegacySlideMasterContentType, System.Text.Encoding.UTF8.GetString(sourceMaster.Root.OuterXml()),
                out LegacyPptImportPart masterAdded, masterRelationshipId);
            PowerPointPartRef targetMaster = PowerPointPartRef.Package(document, masterAdded.PartName!);

            string sourceRelationships = OfficeOpcPartNames.GetRelationshipsPartName(sourceMaster.Name);
            OfficeXmlTree relationships = package.GetXml(sourceRelationships);
            var children = new List<(string Id, string Type, string Target)>();
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                string? target = OfficeOpcPartNames.ResolveRelationshipTarget(sourceRelationships, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                if (target == null || !package.PartExists(target)) continue;
                children.Add((relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id") ?? string.Empty,
                    OfficeXmlNamespaces.NormalizeRelationshipType(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty),
                    target));
            }
            foreach ((string id, string type, string target) in children) {
                string contentType = package.GetContentType(target) ?? throw new InvalidDataException("The scaffold part '" + target + "' has no content type.");
                if (string.Equals(type, LegacySlideLayoutRelationshipType, StringComparison.Ordinal)) {
                    parts.AddXmlPart(targetMaster, type, "/ppt/slideLayouts/slideLayout.xml", contentType,
                        System.Text.Encoding.UTF8.GetString(package.GetXml(target).WriteElement(package.GetXml(target).Root)),
                        out LegacyPptImportPart layoutAdded, id);
                    package.AddRelationship(layoutAdded.PartName, LegacySlideMasterRelationshipType,
                        OfficeOpcPartNames.GetRelativeTarget(layoutAdded.PartName!, targetMaster.Name),
                        id: LegacyPptImportParts.NewSdkRelationshipId());
                } else if (string.Equals(type, LegacyThemeRelationshipType, StringComparison.Ordinal)) {
                    parts.AddXmlPart(targetMaster, type, LegacyPartFolder(targetMaster) + "theme/theme.xml", contentType,
                        System.Text.Encoding.UTF8.GetString(package.GetXml(target).WriteElement(package.GetXml(target).Root)),
                        out _, id);
                } else {
                    throw new InvalidDataException("The native PowerPoint scaffold master has a part the binary import does not copy: " + target);
                }
            }

            OfficeXmlNode presentation = presentationPart.Root;
            OfficeXmlNode masterIds = presentation.GetOrAddChild(LegacyP, "sldMasterIdLst", PPresentationView.Order);
            OfficeXmlNode targetRoot = targetMaster.Root;
            List<OfficeXmlNode> layoutIds = targetRoot.Child(LegacyP, "sldLayoutIdLst").Children(LegacyP, "sldLayoutId").ToList();
            masterIds.AddP("sldMasterId")
                .Set("id", (long)NextLegacyMasterOrLayoutId(presentationPart, excludedLayoutIds: layoutIds))
                .SetRelationship("id", masterRelationshipId);
            // Copied layout ids come from the source and collide with its own.
            foreach (OfficeXmlNode layoutId in layoutIds) layoutId.RemoveAttribute("id");
            foreach (OfficeXmlNode layoutId in layoutIds) {
                layoutId.Set("id", (long)NextLegacyMasterOrLayoutId(presentationPart, excludedLayoutIds: null));
            }
            presentationPart.MarkEdited();
            targetMaster.MarkEdited();
            document.InvalidatePackageRelationships();
            return targetMaster;
        }

        // The next master or layout id: one past the largest id the presentation's masters and their layouts use.
        private static uint NextLegacyMasterOrLayoutId(PowerPointPartRef presentationPart,
            IReadOnlyList<OfficeXmlNode>? excludedLayoutIds) {
            uint max = 2147483648U - 1U;
            foreach (OfficeXmlNode id in presentationPart.Root.Child(LegacyP, "sldMasterIdLst").Children(LegacyP, "sldMasterId").ToList()) {
                if (id.UInt("id") is uint value && value > max) max = value;
            }
            foreach (PowerPointPartRef master in LegacyPptImportParts.Related(presentationPart, LegacySlideMasterRelationshipType)) {
                foreach (OfficeXmlNode id in master.Root.Child(LegacyP, "sldLayoutIdLst")
                             .Children(LegacyP, "sldLayoutId").ToList()) {
                    if (excludedLayoutIds != null && excludedLayoutIds.Contains(id)) continue;
                    if (id.UInt("id") is uint value && value > max) max = value;
                }
            }
            if (max == uint.MaxValue) {
                throw new InvalidOperationException("The presentation has exhausted the slide master and layout identifier range.");
            }
            return max + 1U;
        }

        private static void ApplyLegacyMaster(LegacyPptImportParts parts, PowerPointPartRef masterPart, LegacyPptMaster master,
            LegacyPptHeaderFooterSettings? headerFooter,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            OfficeXmlNode slideMaster = masterPart.Root;
            if (!slideMaster.Is(LegacyP, "sldMaster")) {
                throw new InvalidDataException("The projected PowerPoint package has no slide master.");
            }
            OfficeXmlNode commonSlideData = slideMaster.SetChild(LegacyP, "cSld", LegacyPptImportXml.SlideMasterOrder)
                .Set("name", GetLegacyMasterName(master));
            WriteLegacyShapeTree(commonSlideData.AddP("spTree"), masterPart, master.Shapes,
                master.ConnectorRules, soundContext: soundContext,
                deferredInteractions: deferredInteractions);
            ApplyLegacyMasterRoundTripTheme(parts, masterPart, master.RoundTripTheme);
            if (master.ColorScheme != null
                && master.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacyMasterColorScheme(parts, masterPart, master.ColorScheme);
            }
            if (master.Background != null) {
                ApplyLegacyBackground(masterPart, commonSlideData,
                    master.Background);
            } else if (master.ColorScheme != null) {
                ApplyLegacyBackground(commonSlideData, master.ColorScheme.Background);
            }
            ApplyLegacyMasterTextStyles(masterPart, slideMaster,
                master.TextMasterStyles);
            ApplyLegacyHeaderFooter(slideMaster, commonSlideData,
                headerFooter, allowHeader: false);
            PowerPointPartRef blankLayout = LegacyPptImportParts.Related(masterPart, LegacySlideLayoutRelationshipType).First();
            OfficeXmlNode blankData = blankLayout.Root.Child(LegacyP, "cSld");
            if (!blankData.IsNone) {
                blankData.Set("name", GetLegacyMasterName(master));
                blankLayout.MarkEdited();
            }
            masterPart.MarkEdited();
        }

        private static int AddLegacyTitleMasterLayout(LegacyPptImportParts parts, PowerPointPartRef masterPart,
            LegacyPptMaster titleMaster,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            OfficeOpcEditablePackage package = parts.Package;
            string relationshipId = LegacyPptImportParts.SmallestFreeRelationshipId(package, masterPart.Name);
            parts.AddXmlPart(masterPart, LegacySlideLayoutRelationshipType, "/ppt/slideLayouts/slideLayout.xml",
                LegacySlideLayoutContentType, LegacyP, "sldLayout", _ => { }, out LegacyPptImportPart added, relationshipId);
            PowerPointPartRef layoutPart = PowerPointPartRef.Package(parts.Document, added.PartName!);
            WriteLegacyLayout(layoutPart, GetLegacyMasterName(titleMaster),
                "title", titleMaster.Shapes,
                titleMaster.ConnectorRules, soundContext,
                deferredInteractions);
            package.AddRelationship(layoutPart.Name, LegacySlideMasterRelationshipType, masterPart.RelativeTo(layoutPart),
                id: LegacyPptImportParts.NewSdkRelationshipId());
            parts.Document.InvalidatePackageRelationships();
            OfficeXmlNode layout = layoutPart.Root;
            if (!titleMaster.FollowsMasterObjects) {
                layout.Set("showMasterSp", false);
            }
            ApplyLegacyRoundTripTheme(parts, layoutPart, titleMaster.RoundTripTheme);
            if (!titleMaster.FollowsMasterColorScheme
                && titleMaster.ColorScheme != null
                && titleMaster.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacyColorSchemeOverride(parts, layoutPart, masterPart, titleMaster.ColorScheme);
            }
            OfficeXmlNode commonSlideData = layout.Child(LegacyP, "cSld");
            if (!titleMaster.FollowsMasterBackground && !commonSlideData.IsNone) {
                if (titleMaster.Background != null) {
                    ApplyLegacyBackground(layoutPart,
                        commonSlideData, titleMaster.Background);
                } else if (titleMaster.ColorScheme != null) {
                    ApplyLegacyBackground(commonSlideData,
                        titleMaster.ColorScheme.Background);
                }
            }
            ApplyLegacyHeaderFooter(layout, commonSlideData, titleMaster.HeaderFooter,
                allowHeader: false);

            OfficeXmlNode slideMaster = masterPart.Root;
            if (!slideMaster.Is(LegacyP, "sldMaster")) {
                throw new InvalidDataException("The projected PowerPoint package has no slide master.");
            }
            OfficeXmlNode layoutIds = slideMaster.GetOrAddChild(LegacyP, "sldLayoutIdLst", LegacyPptImportXml.SlideMasterOrder);
            uint layoutId = layoutIds.Children(LegacyP, "sldLayoutId").ToList()
                .Select(item => item.UInt("id") ?? 2147483648U)
                .DefaultIfEmpty(2147483648U)
                .Max() + 1U;
            layoutIds.AddP("sldLayoutId").Set("id", (long)layoutId).SetRelationship("id", relationshipId);
            layoutPart.MarkEdited();
            masterPart.MarkEdited();
            return LegacyPptImportParts.Related(masterPart, LegacySlideLayoutRelationshipType).Count - 1;
        }

        private static string GetLegacyMasterName(LegacyPptMaster master) =>
            $"Binary {(master.IsMainMaster ? "Main" : "Title")} Master {master.MasterId:X8}";

        private static void ProjectLegacySlideDesign(LegacyPptImportParts parts, PowerPointSlide slide, LegacyPptSlide source) {
            PowerPointPartRef slidePart = slide.PartRef;
            OfficeXmlNode slideRoot = slidePart.Root;
            OfficeXmlNode commonSlideData = slideRoot.GetOrAddChild(LegacyP, "cSld", LegacyPptImportXml.SlideOrder);
            ApplyLegacyRoundTripTheme(parts, slidePart,
                source.RoundTripTheme);
            if (!source.FollowsMasterObjects) slideRoot.Set("showMasterSp", false);
            if (!source.FollowsMasterColorScheme && source.ColorScheme != null
                && source.RoundTripTheme?.ThemeXml == null) {
                PowerPointPartRef layoutPart = slidePart.RelatedPart(LegacySlideLayoutRelationshipType);
                ApplyLegacyColorSchemeOverride(parts, slidePart,
                    layoutPart.IsNone ? default : layoutPart.RelatedPart(LegacySlideMasterRelationshipType),
                    source.ColorScheme);
            }
            if (!source.FollowsMasterBackground) {
                if (source.Background != null) {
                    ApplyLegacyBackground(slidePart, commonSlideData,
                        source.Background);
                } else if (source.ColorScheme != null) {
                    ApplyLegacyBackground(commonSlideData, source.ColorScheme.Background);
                }
            }
            slidePart.MarkEdited();
        }

        // The master theme's colour scheme takes the binary colours, as the engine's own theme edits do.
        private static void ApplyLegacyMasterColorScheme(LegacyPptImportParts parts, PowerPointPartRef masterPart,
            LegacyPptColorScheme source) {
            PowerPointPartRef themePart = masterPart.RelatedPart(LegacyThemeRelationshipType);
            if (themePart.IsNone) {
                parts.AddXmlPart(masterPart, LegacyThemeRelationshipType, LegacyPartFolder(masterPart) + "theme/theme.xml",
                    LegacyThemeContentType, LegacyPptXml.A, "theme", theme => theme.AddA("themeElements"), out LegacyPptImportPart added);
                themePart = PowerPointPartRef.Package(parts.Document, added.PartName!);
            }
            OfficeXmlNode theme = themePart.Root;
            OfficeXmlNode elements = theme.GetOrAddChild(LegacyPptXml.A, "themeElements", LegacyPptImportXml.ThemeOrder);
            OfficeXmlNode scheme = elements.Child(LegacyPptXml.A, "clrScheme");
            if (scheme.IsNone) {
                scheme = elements.AddOrdered(LegacyPptXml.A, "clrScheme", LegacyPptImportXml.ThemeElementsOrder).Set("name", "Office");
            }
            SetLegacyThemeColors(scheme, source);
            themePart.MarkEdited();
        }

        /// <summary>
        /// Gives a layout or slide a theme override holding the master's colour scheme with the binary colours.
        /// </summary>
        private static void ApplyLegacyColorSchemeOverride(LegacyPptImportParts parts, PowerPointPartRef target,
            PowerPointPartRef masterPart, LegacyPptColorScheme source) {
            PowerPointPartRef masterTheme = masterPart.IsNone ? default : masterPart.RelatedPart(LegacyThemeRelationshipType);
            OfficeXmlNode masterScheme = masterTheme.IsNone
                ? default
                : masterTheme.Root.Path(LegacyPptXml.A, "themeElements", "clrScheme");
            void Fill(OfficeXmlNode themeOverride) {
                OfficeXmlNode scheme = masterScheme.IsNone
                    ? themeOverride.AddA("clrScheme").Set("name", "Binary PowerPoint")
                    : themeOverride.Append(masterScheme.Clone());
                SetLegacyThemeColors(scheme, source);
            }
            PowerPointPartRef overridePart = target.RelatedPart(LegacyThemeOverrideRelationshipType);
            if (overridePart.IsNone) {
                parts.AddXmlPart(target, LegacyThemeOverrideRelationshipType,
                    LegacyPartFolder(target) + "theme/themeOverride.xml", LegacyThemeOverrideContentType,
                    LegacyPptXml.A, "themeOverride", Fill, out _);
                return;
            }
            OfficeXmlNode existing = overridePart.Root;
            foreach (OfficeXmlNode child in existing.Children().ToList()) child.Remove();
            Fill(existing);
            overridePart.MarkEdited();
        }

        /// <summary>Writes the binary colour scheme into a DrawingML <c>a:clrScheme</c>.</summary>
        internal static void SetLegacyThemeColors(OfficeXmlNode target, LegacyPptColorScheme source) {
            SetLegacyThemeColor(target, "lt1", source.Background);
            SetLegacyThemeColor(target, "dk1", source.Text);
            SetLegacyThemeColor(target, "accent4", source.Shadow);
            SetLegacyThemeColor(target, "dk2", source.TitleText);
            SetLegacyThemeColor(target, "lt2", source.Fill);
            SetLegacyThemeColor(target, "accent1", source.Accent1);
            SetLegacyThemeColor(target, "accent2", source.Accent2);
            SetLegacyThemeColor(target, "accent3", source.Accent3);
        }

        private static void SetLegacyThemeColor(OfficeXmlNode scheme, string name, string value) {
            OfficeXmlNode element = scheme.Child(LegacyPptXml.A, name);
            if (element.IsNone) element = scheme.AddA(name);
            element.RemoveChildren(LegacyPptXml.A, "srgbClr");
            element.RemoveChildren(LegacyPptXml.A, "sysClr");
            element.AddA("srgbClr").Set("val", value);
        }

        /// <summary>Writes a layout root: its type, the shape tree of <paramref name="shapes"/> and the master colour mapping.</summary>
        private static void WriteLegacyLayout(PowerPointPartRef ownerPart, string name, string type,
            IReadOnlyList<LegacyPptShape> shapes,
            IReadOnlyList<LegacyPptConnectorRule>? connectorRules = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions = null) {
            OfficeXmlNode layout = ownerPart.Root;
            ClearLegacyRoot(layout);
            layout.Set("type", type);
            OfficeXmlNode commonSlideData = layout.AddP("cSld").Set("name", name);
            WriteLegacyShapeTree(commonSlideData.AddP("spTree"), ownerPart, shapes,
                connectorRules, soundContext: soundContext,
                deferredInteractions: deferredInteractions);
            layout.AddP("clrMapOvr").AddA("masterClrMapping");
        }

        /// <summary>
        /// Writes an empty <c>p:spTree</c>: the group header, then the projected <paramref name="shapes"/> and connector
        /// rules, adding the parts they need to <paramref name="ownerPart"/>.
        /// </summary>
        private static void WriteLegacyShapeTree(OfficeXmlNode tree, PowerPointPartRef ownerPart = default,
            IReadOnlyList<LegacyPptShape>? shapes = null,
            IReadOnlyList<LegacyPptConnectorRule>? connectorRules = null,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions = null) {
            WriteLegacyShapeTreeHeader(tree);
            if (shapes == null) return;
            if (ownerPart.IsNone) {
                throw new InvalidOperationException("An owning Open XML part is required to project legacy shapes.");
            }

            uint nextShapeId = 2U;
            var projectedShapeIds = new Dictionary<uint, uint>();
            foreach (LegacyPptShape source in shapes) {
                CreateLegacyOpenXmlShape(ownerPart, tree, source,
                    ref nextShapeId, slidePartsByLegacyId, soundContext,
                    projectedShapeIds, deferredInteractions);
            }
            if (connectorRules != null) {
                ProjectLegacyConnectorRules(tree, connectorRules, projectedShapeIds);
            }
        }

        /// <summary>
        /// Projects a binary shape under <paramref name="parent"/> (a shape tree or group) and returns the element added,
        /// or none when the shape has no native form.
        /// </summary>
        internal static OfficeXmlNode CreateLegacyOpenXmlShape(PowerPointPartRef ownerPart,
            OfficeXmlNode parent, LegacyPptShape source, ref uint nextShapeId,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            IDictionary<uint, uint>? projectedShapeIds = null,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions = null) {
            if (ownerPart.IsNone) throw new ArgumentNullException(nameof(ownerPart));
            if (parent.IsNone) throw new ArgumentNullException(nameof(parent));
            ArgumentNullException.ThrowIfNull(source);
            uint shapeId = nextShapeId++;
            OfficeXmlNode projected;
            switch (source.Kind) {
                case LegacyPptShapeKind.Picture:
                    projected = CreateLegacyPicture(ownerPart, parent, source, shapeId);
                    break;
                case LegacyPptShapeKind.Table:
                    if (ownerPart.Document is { IsSdkLoaded: false } document) {
                        // The table projection builds through the SDK. While the masters still add parts through the
                        // package, a placeholder holds the table's place until the slides are projected.
                        projected = AddLegacyLateTable(LegacyPptImportParts.For(document), ownerPart, parent, source, shapeId);
                        RegisterLegacyShapeId(source, projected, projectedShapeIds);
                        return projected;
                    }
                    projected = parent.Append(CreateLegacyTableFrameElement(ownerPart,
                        source, shapeId, slidePartsByLegacyId, soundContext,
                        deferredInteractions));
                    break;
                case LegacyPptShapeKind.Group:
                    projected = CreateLegacyGroupShape(ownerPart, parent, source, shapeId,
                        ref nextShapeId, slidePartsByLegacyId, soundContext,
                        projectedShapeIds, deferredInteractions);
                    break;
                default:
                    projected = CreateLegacyShape(ownerPart, parent, source, shapeId,
                        slidePartsByLegacyId, soundContext,
                        deferredInteractions);
                    break;
            }
            if (!projected.IsNone) {
                RegisterLegacyShapeId(source, projected,
                    projectedShapeIds);
                ApplyLegacyShapeMetadata(projected, source);
                ApplyLegacyShapeInteractions(ownerPart, projected, source,
                    slidePartsByLegacyId, soundContext,
                    deferredInteractions);
            }
            return projected;
        }

        internal static bool CanCreateLegacyOpenXmlShape(
            LegacyPptShape source) {
            ArgumentNullException.ThrowIfNull(source);
            return source.Kind switch {
                LegacyPptShapeKind.TextBox or
                LegacyPptShapeKind.Rectangle or
                LegacyPptShapeKind.Ellipse or
                LegacyPptShapeKind.Line => true,
                LegacyPptShapeKind.AutoShape or
                LegacyPptShapeKind.Connector =>
                    LegacyPptShapeGeometryMapper.TryGetPreset(
                        source.OfficeArtShapeType, out _),
                LegacyPptShapeKind.Picture =>
                    source.Picture?.HasImportableImage == true
                    && source.Picture.ContentType != null,
                LegacyPptShapeKind.Table => source.Table != null,
                LegacyPptShapeKind.Group =>
                    source.GroupCoordinateBounds.HasValue
                    && source.Children.Count > 0,
                _ => false
            };
        }

        private static void ApplyLegacyShapeMetadata(OfficeXmlNode target, LegacyPptShape source) {
            OfficeXmlNode properties = GetLegacyNonVisualDrawingProperties(target);
            if (properties.IsNone) return;
            if (source.Metadata.Name != null) properties.Set("name", source.Metadata.Name);
            if (source.Metadata.Description != null) properties.Set("descr", source.Metadata.Description);
            properties.Set("hidden", source.Style.Hidden == true ? true : null);
        }

        private static OfficeXmlNode CreateLegacyShape(PowerPointPartRef ownerPart, OfficeXmlNode parent,
            LegacyPptShape source, uint shapeId,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions) {
            if (source.Kind == LegacyPptShapeKind.Connector
                && LegacyPptShapeGeometryMapper.TryGetPreset(source.OfficeArtShapeType,
                    out OfficeIMO.Drawing.OfficePresetShapeType connectorGeometry)) {
                return CreateLegacyConnectionShape(parent, source, shapeId, connectorGeometry.ToOpenXmlToken());
            }
            string geometry;
            switch (source.Kind) {
                case LegacyPptShapeKind.TextBox:
                    geometry = source.OfficeArtShapeType != 202
                        && LegacyPptShapeGeometryMapper.TryGetPreset(source.OfficeArtShapeType,
                            out OfficeIMO.Drawing.OfficePresetShapeType textGeometry)
                        ? textGeometry.ToOpenXmlToken()
                        : "rect";
                    break;
                case LegacyPptShapeKind.Rectangle:
                    geometry = "rect";
                    break;
                case LegacyPptShapeKind.Ellipse:
                    geometry = "ellipse";
                    break;
                case LegacyPptShapeKind.Line:
                    geometry = "line";
                    break;
                case LegacyPptShapeKind.AutoShape:
                    if (!LegacyPptShapeGeometryMapper.TryGetPreset(source.OfficeArtShapeType,
                            out OfficeIMO.Drawing.OfficePresetShapeType autoGeometry)) return default;
                    geometry = autoGeometry.ToOpenXmlToken();
                    break;
                default:
                    return default;
            }

            OfficeXmlNode shape = parent.AddP("sp");
            OfficeXmlNode nonVisual = shape.AddP("nvSpPr");
            nonVisual.AddP("cNvPr").Set("id", (long)shapeId).Set("name", $"Binary Shape {shapeId - 1U}");
            nonVisual.AddP("cNvSpPr").AddA("spLocks").Set("noGrp", true);
            OfficeXmlNode applicationProperties = nonVisual.AddP("nvPr");
            if (source.Placeholder != null) {
                WriteLegacyPlaceholderShape(applicationProperties.AddP("ph"), source.Placeholder);
            }
            OfficeXmlNode shapeProperties = shape.AddP("spPr");
            OfficeXmlNode transform = shapeProperties.AddA("xfrm");
            AppendLegacyBoundsTransform(transform, source);
            OfficeXmlNode preset = shapeProperties.AddA("prstGeom").Set("prst", geometry);
            preset.AddA("avLst");
            if (geometry == "line") shapeProperties.AddA("noFill");
            LegacyPptShapeGeometryMapper.ApplyExactPresetAdjustments(source.OfficeArtShapeType,
                source.Geometry, preset);
            ApplyLegacyShapeStyle(shapeProperties, source);
            ApplyLegacyShapeTransform(transform, source);

            if (source.Kind == LegacyPptShapeKind.TextBox) {
                bool hasRichText = LegacyPptTextProjection.HasRichText(source.TextBody);
                bool deferTextInteractions =
                    ShouldDeferLegacyTextInteractions(source.TextBody,
                        slidePartsByLegacyId, deferredInteractions);
                OfficeXmlNode textBody = shape.AddP("txBody");
                if (hasRichText) {
                    LegacyPptTextProjection.WriteTextBody(textBody, source.TextBody,
                        source.TextFrame,
                        deferTextInteractions
                            ? SuppressedLegacyTextInteractions
                            : CreateLegacyTextInteractionProjector(ownerPart,
                                slidePartsByLegacyId, soundContext),
                        pictureBullet => ProjectLegacyPictureBullet(ownerPart,
                            pictureBullet));
                } else {
                    OfficeXmlNode bodyProperties = textBody.AddA("bodyPr");
                    textBody.AddA("lstStyle");
                    foreach (string paragraph in PowerPointTextContent.SplitParagraphs(source.Text)) {
                        PowerPointTextContent.InsertInlineContent(textBody.AddA("p"), paragraph, default, default);
                    }
                    LegacyPptTextProjection.ApplyTextFrame(bodyProperties, source.TextFrame);
                }
                if (deferTextInteractions) {
                    deferredInteractions!.Add(new LegacyPptDeferredProjection(
                        projectedSlides => {
                            OfficeXmlNode current = FindLegacyShapeById(ownerPart.Root, shapeId);
                            if (current.IsNone) return;
                            LegacyPptTextProjection.WriteTextBody(
                                current.SetChild(LegacyP, "txBody", LegacyPptImportXml.ShapeOrder),
                                source.TextBody,
                                source.TextFrame,
                                CreateLegacyTextInteractionProjector(ownerPart,
                                    projectedSlides, soundContext),
                                pictureBullet =>
                                    ProjectLegacyPictureBullet(ownerPart,
                                        pictureBullet));
                            ownerPart.MarkEdited();
                        }));
                }
            }
            return shape;
        }

        // The offset and extents of a shape's bounds, at least one EMU wide and high.
        private static void AppendLegacyBoundsTransform(OfficeXmlNode transform, LegacyPptShape source) =>
            AppendLegacyOffsetAndExtents(transform, ToEmus(source.Bounds.Left), ToEmus(source.Bounds.Top),
                Math.Max(1L, ToEmus(source.Bounds.Width)), Math.Max(1L, ToEmus(source.Bounds.Height)));

        private static OfficeXmlNode CreateLegacyConnectionShape(OfficeXmlNode parent, LegacyPptShape source, uint shapeId,
            string geometry) {
            OfficeXmlNode connector = parent.AddP("cxnSp");
            OfficeXmlNode nonVisual = connector.AddP("nvCxnSpPr");
            nonVisual.AddP("cNvPr").Set("id", (long)shapeId).Set("name", $"Binary Connector {shapeId - 1U}");
            nonVisual.AddP("cNvCxnSpPr");
            nonVisual.AddP("nvPr");
            OfficeXmlNode properties = connector.AddP("spPr");
            OfficeXmlNode transform = properties.AddA("xfrm");
            AppendLegacyBoundsTransform(transform, source);
            OfficeXmlNode preset = properties.AddA("prstGeom").Set("prst", geometry);
            preset.AddA("avLst");
            LegacyPptShapeGeometryMapper.ApplyExactPresetAdjustments(source.OfficeArtShapeType,
                source.Geometry, preset);
            ApplyLegacyShapeStyle(properties, source);
            ApplyLegacyShapeTransform(transform, source);
            return connector;
        }

        private static OfficeXmlNode CreateLegacyGroupShape(PowerPointPartRef ownerPart, OfficeXmlNode parent, LegacyPptShape source,
            uint shapeId, ref uint nextShapeId,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext,
            IDictionary<uint, uint>? projectedShapeIds,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions) {
            if (!source.GroupCoordinateBounds.HasValue || source.Children.Count == 0) return default;
            LegacyPptBounds coordinate = source.GroupCoordinateBounds.Value;
            OfficeXmlNode group = parent.AddP("grpSp");
            OfficeXmlNode nonVisual = group.AddP("nvGrpSpPr");
            nonVisual.AddP("cNvPr").Set("id", (long)shapeId).Set("name", $"Binary Group {shapeId - 1U}");
            nonVisual.AddP("cNvGrpSpPr");
            nonVisual.AddP("nvPr");
            OfficeXmlNode transform = group.AddP("grpSpPr").AddA("xfrm");
            AppendLegacyBoundsTransform(transform, source);
            transform.AddA("chOff").Set("x", ToEmus(coordinate.Left)).Set("y", ToEmus(coordinate.Top));
            transform.AddA("chExt")
                .Set("cx", Math.Max(1L, ToEmus(coordinate.Width)))
                .Set("cy", Math.Max(1L, ToEmus(coordinate.Height)));
            ApplyLegacyShapeTransform(transform, source);
            foreach (LegacyPptShape child in source.Children) {
                CreateLegacyOpenXmlShape(ownerPart, group, child,
                    ref nextShapeId, slidePartsByLegacyId, soundContext,
                    projectedShapeIds, deferredInteractions);
            }
            return group;
        }

        private static OfficeXmlNode CreateLegacyPicture(PowerPointPartRef ownerPart, OfficeXmlNode parent, LegacyPptShape source,
            uint shapeId) {
            if (source.Picture?.HasImportableImage != true || source.Picture.ContentType == null) return default;
            string relationshipId = AddLegacyImagePart(ownerPart, source.Picture);
            OfficeXmlNode picture = parent.AddP("pic");
            OfficeXmlNode nonVisual = picture.AddP("nvPicPr");
            nonVisual.AddP("cNvPr").Set("id", (long)shapeId).Set("name", $"Binary Picture {shapeId - 1U}");
            WriteLegacyNonVisualPictureDrawingProperties(nonVisual.AddP("cNvPicPr"), source);
            nonVisual.AddP("nvPr");
            OfficeXmlNode blipFill = picture.AddP("blipFill");
            OfficeXmlNode blip = blipFill.AddA("blip").SetRelationship("embed", relationshipId);
            blipFill.AddA("stretch").AddA("fillRect");
            ApplyLegacyPictureCrop(blipFill, source);
            ApplyLegacyPictureEffects(blip, source);
            OfficeXmlNode properties = picture.AddP("spPr");
            OfficeXmlNode transform = properties.AddA("xfrm");
            AppendLegacyBoundsTransform(transform, source);
            properties.AddA("prstGeom").Set("prst", "rect").AddA("avLst");
            ApplyLegacyShapeTransform(transform, source);
            return picture;
        }

        /// <summary>Relates the owner to the BLIP's image, reusing an identical image the import already gave it.</summary>
        private static string AddLegacyImagePart(PowerPointPartRef ownerPart,
            OfficeIMO.Drawing.Binary.OfficeArtBlipStoreEntry picture) {
            if (picture.ContentType == null || !picture.HasImportableImage) {
                throw new ArgumentException("The OfficeArt BLIP has no importable image payload.",
                    nameof(picture));
            }
            GetLegacyPicturePartType(picture.ContentType);
            return LegacyPptImportParts.For(RequireLegacyDocument(ownerPart)).AddImage(ownerPart,
                picture.ContentType, picture.ImageBytes, reuseExisting: true);
        }

        private static PowerPointDocument RequireLegacyDocument(PowerPointPartRef part) =>
            part.Document ?? throw new InvalidOperationException("The projected part does not belong to a presentation.");

        internal static string? ProjectLegacyPictureBullet(
            PowerPointPartRef ownerPart, LegacyPptPictureBullet pictureBullet) {
            if (!pictureBullet.HasImportableImage
                || pictureBullet.ContentType == null) return null;
            LegacyPptImportParts parts = LegacyPptImportParts.For(RequireLegacyDocument(ownerPart));
            Dictionary<ushort, string> relationships = parts.PictureBulletsOf(ownerPart);
            if (relationships.TryGetValue(pictureBullet.Index,
                    out string? existingRelationship)) {
                return existingRelationship;
            }
            GetLegacyPicturePartType(pictureBullet.ContentType);
            string relationshipId = parts.AddImage(ownerPart,
                pictureBullet.ContentType, pictureBullet.ImageBytes,
                reuseExisting: false);
            relationships.Add(pictureBullet.Index, relationshipId);
            return relationshipId;
        }

        private static void ApplyLegacyPictureCrop(OfficeXmlNode target, LegacyPptShape source) {
            if (target.IsNone || !source.PictureProperties.HasCrop) return;
            target.SetChild(LegacyPptXml.A, "srcRect", LegacyPptImportXml.BlipFillOrder)
                .Set("l", ToOpenXmlCrop(source.PictureProperties.CropFromLeft))
                .Set("t", ToOpenXmlCrop(source.PictureProperties.CropFromTop))
                .Set("r", ToOpenXmlCrop(source.PictureProperties.CropFromRight))
                .Set("b", ToOpenXmlCrop(source.PictureProperties.CropFromBottom));
        }

        private static int? ToOpenXmlCrop(double? fraction) {
            if (!fraction.HasValue || double.IsNaN(fraction.Value)
                || double.IsInfinity(fraction.Value)) return null;
            double value = Math.Round(fraction.Value * 100000D, MidpointRounding.AwayFromZero);
            return value < int.MinValue || value > int.MaxValue ? null : (int)value;
        }

        private static void ApplyLegacyPictureEffects(OfficeXmlNode target, LegacyPptShape source) {
            if (target.IsNone) return;
            if (source.PictureTransparentColor != null) {
                OfficeXmlNode change = target.AddA("clrChange");
                change.AddA("clrFrom").AddA("srgbClr").Set("val", source.PictureTransparentColor);
                change.AddA("clrTo").AddA("srgbClr").Set("val", source.PictureTransparentColor)
                    .AddA("alpha").Set("val", 0L);
            }
            if (source.PictureRecolorColor != null) {
                target.AddA("clrRepl").AddA("srgbClr").Set("val", source.PictureRecolorColor);
            }
            int? brightness = ToOpenXmlPictureAdjustment(
                source.PictureProperties.BrightnessAdjustment);
            int? contrast = ToOpenXmlPictureAdjustment(
                source.PictureProperties.ContrastAdjustment);
            if (brightness.HasValue || contrast.HasValue) {
                target.AddA("lum").Set("bright", brightness).Set("contrast", contrast);
            }
            if (source.PictureProperties.BiLevel == true) {
                target.AddA("biLevel").Set("thresh", 50000L);
            } else if (source.PictureProperties.Grayscale == true) {
                target.AddA("grayscl");
            }
        }

        private static int? ToOpenXmlPictureAdjustment(double? fraction) {
            if (!fraction.HasValue || double.IsNaN(fraction.Value)
                || double.IsInfinity(fraction.Value)) return null;
            return (int)Math.Round(Math.Max(-1D, Math.Min(1D, fraction.Value)) * 100000D,
                MidpointRounding.AwayFromZero);
        }

        /// <summary>Writes the binary rotation and flips onto <c>a:xfrm</c> of a shape or group.</summary>
        internal static void ApplyLegacyShapeTransform(OfficeXmlNode transform, LegacyPptShape source) {
            transform.Set("rot", ToOpenXmlRotation(source.Transform.RotationDegrees));
            transform.Set("flipH", source.Transform.FlipHorizontal ? true : null);
            transform.Set("flipV", source.Transform.FlipVertical ? true : null);
        }

        private static int? ToOpenXmlRotation(double? degrees) {
            if (!degrees.HasValue) return null;
            double value = degrees.Value * 60000D;
            if (value < int.MinValue || value > int.MaxValue) return null;
            return (int)Math.Round(value);
        }

        /// <summary>Writes the binary fill, line and shadow into shape properties (<c>p:spPr</c>).</summary>
        internal static void ApplyLegacyShapeStyle(OfficeXmlNode properties, LegacyPptShape source) {
            OfficeIMO.Drawing.Binary.OfficeArtShapeStyle style = source.Style;
            if (style.FillEnabled == false) {
                SetLegacyShapeFill(properties, properties.CreateElement(LegacyPptXml.A, "noFill"));
            } else if (CreateLegacyShapeGradientFill(properties, source) is { IsNone: false } gradient) {
                SetLegacyShapeFill(properties, gradient);
            } else if (source.FillColor != null && style.FillType.GetValueOrDefault() == 0) {
                SetLegacyShapeFill(properties, CreateLegacySolidFill(properties, source.FillColor, style.FillOpacity));
            }
            bool hasLineStyle = style.LineEnabled.HasValue || source.LineColor != null
                || style.LineOpacity.HasValue || style.LineWidthEmus.HasValue || style.LineDashing.HasValue
                || style.LineStartArrowhead.HasValue || style.LineEndArrowhead.HasValue
                || style.LineJoinStyle.HasValue || style.LineEndCapStyle.HasValue;
            if (!hasLineStyle) {
                ApplyLegacyShapeShadow(properties, source);
                return;
            }

            OfficeXmlNode outline = properties.Child(LegacyPptXml.A, "ln");
            if (outline.IsNone) outline = properties.AddA("ln");
            if (style.LineEnabled == false) {
                SetLegacyOutlineFill(outline, outline.CreateElement(LegacyPptXml.A, "noFill"));
                ApplyLegacyShapeShadow(properties, source);
                return;
            }

            if (source.LineColor != null) {
                SetLegacyOutlineFill(outline, CreateLegacySolidFill(outline, source.LineColor, style.LineOpacity));
            }
            if (style.LineWidthEmus is >= 0) outline.Set("w", (long)style.LineWidthEmus.Value);
            ApplyLegacyLineDash(outline, style.LineDashing);
            ApplyLegacyLineJoin(outline, style.LineJoinStyle);
            outline.Set("cap", MapLegacyLineCap(style.LineEndCapStyle));
            ApplyLegacyLineEnd(outline, isHead: true, style.LineStartArrowhead,
                style.LineStartArrowWidth, style.LineStartArrowLength);
            ApplyLegacyLineEnd(outline, isHead: false, style.LineEndArrowhead,
                style.LineEndArrowWidth, style.LineEndArrowLength);
            ApplyLegacyShapeShadow(properties, source);
        }

        private static readonly HashSet<string> LegacyShadowPredecessors = new(StringComparer.Ordinal) {
            "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln"
        };

        private static void ApplyLegacyShapeShadow(OfficeXmlNode properties,
            LegacyPptShape source) {
            OfficeIMO.Drawing.Binary.OfficeArtShapeStyle style = source.Style;
            if (!style.HasProjectableShadow) return;

            int offsetX = style.ShadowOffsetXEmus ?? 0x6338;
            int offsetY = style.ShadowOffsetYEmus ?? 0x6338;
            double angle = Math.Atan2(offsetY, offsetX) * 180D / Math.PI;
            if (angle < 0D) angle += 360D;
            long distance = (long)Math.Round(Math.Sqrt((double)offsetX * offsetX
                + (double)offsetY * offsetY), MidpointRounding.AwayFromZero);
            OfficeXmlNode effects = properties.Child(LegacyPptXml.A, "effectLst");
            bool added = effects.IsNone;
            if (added) effects = properties.CreateElement(LegacyPptXml.A, "effectLst");
            effects.RemoveChildren(LegacyPptXml.A, "outerShdw");
            OfficeXmlNode shadow = effects.AddA("outerShdw")
                .Set("blurRad", (long)Math.Max(0, style.ShadowSoftnessEmus ?? 0))
                .Set("dist", distance)
                .Set("dir", (long)(int)Math.Round(angle * 60000D, MidpointRounding.AwayFromZero))
                .Set("rotWithShape", false);
            shadow.AddA("srgbClr").Set("val", source.ShadowColor ?? "808080")
                .AddA("alpha").Set("val", (long)checked((int)Math.Round(
                    Math.Max(0D, Math.Min(1D, style.ShadowOpacity ?? 1D)) * 100000D)));
            if (added) {
                OfficeXmlNode insertBefore = default;
                foreach (OfficeXmlNode child in properties.Children()) {
                    if (child.NamespaceId == LegacyPptXml.A && LegacyShadowPredecessors.Contains(child.LocalName)) continue;
                    insertBefore = child;
                    break;
                }
                properties.InsertBefore(effects, insertBefore);
            }
        }

        private static OfficeXmlNode CreateLegacySolidFill(OfficeXmlNode context, string color, double? opacity) {
            OfficeXmlNode fill = context.CreateElement(LegacyPptXml.A, "solidFill");
            OfficeXmlNode rgb = fill.AddA("srgbClr").Set("val", color);
            if (opacity.HasValue) {
                rgb.AddA("alpha").Set("val", (long)checked((int)Math.Round(
                    Math.Max(0D, Math.Min(1D, opacity.Value)) * 100000D)));
            }
            return fill;
        }

        private static void SetLegacyShapeFill(OfficeXmlNode properties, OfficeXmlNode fill) {
            properties.RemoveChildren(LegacyPptXml.A, "noFill");
            properties.RemoveChildren(LegacyPptXml.A, "solidFill");
            properties.RemoveChildren(LegacyPptXml.A, "gradFill");
            properties.RemoveChildren(LegacyPptXml.A, "blipFill");
            properties.RemoveChildren(LegacyPptXml.A, "pattFill");
            properties.InsertBefore(fill, properties.Child(LegacyPptXml.A, "ln"));
        }

        private static void SetLegacyOutlineFill(OfficeXmlNode outline, OfficeXmlNode fill) {
            outline.RemoveChildren(LegacyPptXml.A, "noFill");
            outline.RemoveChildren(LegacyPptXml.A, "solidFill");
            outline.RemoveChildren(LegacyPptXml.A, "gradFill");
            outline.RemoveChildren(LegacyPptXml.A, "pattFill");
            outline.InsertBefore(fill, outline.FirstChild);
        }

        private static void ApplyLegacyLineDash(OfficeXmlNode outline, uint? value) {
            outline.RemoveChildren(LegacyPptXml.A, "prstDash");
            string? dash = value switch {
                0 => "solid",
                1 => "sysDash",
                2 => "sysDot",
                3 => "sysDashDot",
                4 => "sysDashDotDot",
                5 => "dot",
                6 => "dash",
                7 => "lgDash",
                8 => "dashDot",
                9 => "lgDashDot",
                10 => "lgDashDotDot",
                _ => null
            };
            if (dash != null) InsertLegacyOutlineChild(outline, outline.CreateElement(LegacyPptXml.A, "prstDash").Set("val", dash));
        }

        private static void ApplyLegacyLineJoin(OfficeXmlNode outline, uint? value) {
            outline.RemoveChildren(LegacyPptXml.A, "bevel");
            outline.RemoveChildren(LegacyPptXml.A, "miter");
            outline.RemoveChildren(LegacyPptXml.A, "round");
            string? join = value switch {
                0 => "bevel",
                1 => "miter",
                2 => "round",
                _ => null
            };
            if (join != null) InsertLegacyOutlineChild(outline, outline.CreateElement(LegacyPptXml.A, join));
        }

        private static string? MapLegacyLineCap(uint? value) => value switch {
            0 => "rnd",
            1 => "sq",
            2 => "flat",
            _ => null
        };

        private static void ApplyLegacyLineEnd(OfficeXmlNode outline, bool isHead, uint? type,
            uint? width, uint? length) {
            if (!type.HasValue && !width.HasValue && !length.HasValue) return;
            string name = isHead ? "headEnd" : "tailEnd";
            OfficeXmlNode end = outline.Child(LegacyPptXml.A, name);
            bool added = end.IsNone;
            if (added) end = outline.CreateElement(LegacyPptXml.A, name);
            end.Set("type", MapLegacyLineEnd(type) ?? "none")
                .Set("w", MapLegacyLineEndSize(width))
                .Set("len", MapLegacyLineEndSize(length));
            if (added) InsertLegacyOutlineChild(outline, end);
        }

        private static string? MapLegacyLineEnd(uint? value) => value switch {
            0 => "none",
            1 => "triangle",
            2 => "stealth",
            3 => "diamond",
            4 => "oval",
            5 => "arrow",
            _ => null
        };

        private static string? MapLegacyLineEndSize(uint? value) => value switch {
            0 => "sm",
            1 => "med",
            2 => "lg",
            _ => null
        };

        private static int GetLegacyOutlineChildRank(OfficeXmlNode child) =>
            child.NamespaceId != LegacyPptXml.A ? 100 : child.LocalName switch {
                "noFill" or "solidFill" or "gradFill" or "pattFill" => 0,
                "prstDash" or "custDash" => 1,
                "round" or "bevel" or "miter" => 2,
                "headEnd" => 3,
                "tailEnd" => 4,
                _ => 100
            };

        private static void InsertLegacyOutlineChild(OfficeXmlNode outline, OfficeXmlNode child) {
            int order = GetLegacyOutlineChildRank(child);
            OfficeXmlNode before = default;
            foreach (OfficeXmlNode existing in outline.Children()) {
                if (GetLegacyOutlineChildRank(existing) > order) {
                    before = existing;
                    break;
                }
            }
            outline.InsertBefore(child, before);
        }
    }
}
