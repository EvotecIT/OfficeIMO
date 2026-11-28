using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ProjectLegacySpecialMasters(LegacyPptImportParts parts,
            LegacyPptDocument legacy,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            if (legacy.NotesMaster != null) {
                ProjectLegacyNotesMaster(parts, legacy.NotesMaster,
                    legacy.NotesHeaderFooterDefaults, soundContext,
                    deferredInteractions);
            }
            if (legacy.HandoutMaster != null) {
                ProjectLegacyHandoutMaster(parts, legacy.HandoutMaster,
                    legacy.NotesHeaderFooterDefaults, soundContext,
                    deferredInteractions);
            }
        }

        private static void ProjectLegacyNotesMaster(LegacyPptImportParts parts,
            LegacyPptSpecialMaster source,
            LegacyPptHeaderFooterSettings? headerFooter,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            PowerPointPartRef presentationPart = parts.PresentationPart;
            PowerPointPartRef part = presentationPart.RelatedPart(LegacyNotesMasterRelationshipType);
            OfficeXmlNode target = part.IsNone ? default : part.Root;
            if (target.IsNone || !target.Is(LegacyPptXml.P, "notesMaster")
                || part.RelatedPart(LegacyThemeRelationshipType).IsNone) {
                throw new InvalidDataException("The native PowerPoint scaffold has no notes master.");
            }
            OfficeXmlNode commonSlideData = target.SetChild(LegacyPptXml.P, "cSld", LegacyPptImportXml.NotesMasterOrder);
            WriteLegacySpecialMasterCommonSlideData(part, commonSlideData, source,
                "Binary Notes Master", soundContext, deferredInteractions);
            ApplyLegacyMasterRoundTripTheme(parts, part, source.RoundTripTheme);
            if (source.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacySpecialMasterTheme(part.RelatedPart(LegacyThemeRelationshipType),
                    source.ColorScheme);
            }
            ApplyLegacyHeaderFooter(target, commonSlideData,
                headerFooter, allowHeader: true);
            part.MarkEdited();
        }

        private static void ProjectLegacyHandoutMaster(LegacyPptImportParts parts,
            LegacyPptSpecialMaster source,
            LegacyPptHeaderFooterSettings? headerFooter,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            PowerPointPartRef presentationPart = parts.PresentationPart;
            PowerPointPartRef part = presentationPart.RelatedPart(LegacyHandoutMasterRelationshipType);
            OfficeXmlNode handoutMaster;
            if (part.IsNone) {
                parts.AddXmlPart(presentationPart, LegacyHandoutMasterRelationshipType,
                    "/ppt/handoutMasters/handoutMaster.xml", LegacyHandoutMasterContentType,
                    LegacyPptXml.P, "handoutMaster", _ => { }, out LegacyPptImportPart added);
                part = PowerPointPartRef.Package(parts.Document, added.PartName!);
            }
            EnsureLegacySpecialMasterTheme(parts, part, presentationPart);

            handoutMaster = part.Root;
            foreach (OfficeXmlNode child in handoutMaster.Children().ToList()) child.Remove();
            OfficeXmlNode commonSlideData = handoutMaster.AddP("cSld");
            WriteLegacySpecialMasterCommonSlideData(part, commonSlideData,
                source, "Binary Handout Master", soundContext,
                deferredInteractions);
            OfficeXmlNode colorMap = handoutMaster.AddP("clrMap");
            WriteLegacyColorMap(colorMap, presentationPart);
            ApplyLegacyMasterRoundTripTheme(parts, part, source.RoundTripTheme);
            if (source.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacySpecialMasterTheme(part.RelatedPart(LegacyThemeRelationshipType),
                    source.ColorScheme);
            }
            ApplyLegacyHeaderFooter(handoutMaster, commonSlideData,
                headerFooter, allowHeader: true);
            part.MarkEdited();

            OfficeXmlNode root = presentationPart.Root;
            OfficeXmlNode list = root.GetOrAddChild(LegacyPptXml.P, "handoutMasterIdLst", PPresentationView.Order);
            string? relationshipId = parts.RelationshipIdTo(presentationPart, part);
            if (!list.Children(LegacyPptXml.P, "handoutMasterId").ToList().Any(item =>
                    item.RelationshipId() == relationshipId)) {
                list.AddP("handoutMasterId").SetRelationship("id", relationshipId);
            }
            presentationPart.MarkEdited();
        }

        private static void WriteLegacySpecialMasterCommonSlideData(
            PowerPointPartRef ownerPart, OfficeXmlNode commonSlideData, LegacyPptSpecialMaster source, string name,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            commonSlideData.Set("name", name);
            WriteLegacyShapeTree(commonSlideData.AddP("spTree"), ownerPart, source.Shapes,
                source.ConnectorRules, soundContext: soundContext,
                deferredInteractions: deferredInteractions);
            if (source.Background != null) {
                ApplyLegacyBackground(ownerPart, commonSlideData, source.Background);
            } else if (source.ColorScheme != null) {
                ApplyLegacyBackground(commonSlideData, source.ColorScheme.Background);
            }
        }

        private static void EnsureLegacySpecialMasterTheme(LegacyPptImportParts parts, PowerPointPartRef target,
            PowerPointPartRef presentationPart) {
            if (!target.RelatedPart(LegacyThemeRelationshipType).IsNone) return;
            PowerPointPartRef firstMaster = presentationPart.RelatedPart(LegacySlideMasterRelationshipType);
            PowerPointPartRef sourceTheme = firstMaster.IsNone ? default : firstMaster.RelatedPart(LegacyThemeRelationshipType);
            OfficeXmlNode source = sourceTheme.IsNone ? default : sourceTheme.Root;
            string themePartName = LegacyPartFolder(target) + "theme/theme.xml";
            if (!source.IsNone && source.Is(LegacyPptXml.A, "theme")) {
                // A copy of the whole theme, its namespace declarations included.
                parts.AddXmlPart(target, LegacyThemeRelationshipType, themePartName, LegacyThemeContentType,
                    System.Text.Encoding.UTF8.GetString(source.OuterXml()), out _);
                return;
            }
            parts.AddXmlPart(target, LegacyThemeRelationshipType, themePartName,
                LegacyThemeContentType, LegacyPptXml.A, "theme", theme => {
                    theme.Set("name", "Binary PowerPoint");
                    theme.AddA("themeElements");
                }, out _);
        }

        private static void ApplyLegacySpecialMasterTheme(PowerPointPartRef themePart,
            LegacyPptColorScheme? source) {
            if (themePart.IsNone || source == null) return;
            OfficeXmlNode theme = themePart.Root;
            if (theme.IsNone) return;
            OfficeXmlNode elements = theme.GetOrAddChild(LegacyPptXml.A, "themeElements", LegacyPptImportXml.ThemeOrder);
            OfficeXmlNode target = elements.Child(LegacyPptXml.A, "clrScheme");
            if (target.IsNone) {
                target = elements.AddOrdered(LegacyPptXml.A, "clrScheme", LegacyPptImportXml.ThemeElementsOrder)
                    .Set("name", "Binary PowerPoint");
            }
            SetLegacyThemeColors(target, source);
            themePart.MarkEdited();
        }

        // The first slide master's colour map, or the default mapping when it has none.
        private static void WriteLegacyColorMap(OfficeXmlNode target, PowerPointPartRef presentationPart) {
            PowerPointPartRef firstMaster = presentationPart.RelatedPart(LegacySlideMasterRelationshipType);
            OfficeXmlNode source = firstMaster.IsNone ? default : firstMaster.Root.Child(LegacyPptXml.P, "clrMap");
            if (!source.IsNone) {
                CopyLegacyAttributes(source, target);
                target.AppendCopiesOfChildren(source);
                return;
            }
            target.Set("bg1", "lt1").Set("tx1", "dk1").Set("bg2", "lt2").Set("tx2", "dk2")
                .Set("accent1", "accent1").Set("accent2", "accent2").Set("accent3", "accent3")
                .Set("accent4", "accent4").Set("accent5", "accent5").Set("accent6", "accent6")
                .Set("hlink", "hlink").Set("folHlink", "folHlink");
        }
    }
}
