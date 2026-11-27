using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ApplyLegacyMasterTextStyles(
            PowerPointPartRef ownerPart, OfficeXmlNode slideMaster,
            IReadOnlyList<LegacyPptTextMasterStyle> source) {
            if (source.Count == 0) return;
            LegacyPptTextMasterStyle? title = FindUsableStyle(source,
                LegacyPptTextType.Title, LegacyPptTextType.CenterTitle);
            LegacyPptTextMasterStyle? body = FindUsableStyle(source,
                LegacyPptTextType.Body, LegacyPptTextType.CenterBody,
                LegacyPptTextType.HalfBody, LegacyPptTextType.QuarterBody);
            LegacyPptTextMasterStyle? other = FindUsableStyle(source,
                LegacyPptTextType.Other, LegacyPptTextType.Notes);

            OfficeXmlNode styles = slideMaster.SetChild(LegacyPptXml.P, "txStyles", LegacyPptImportXml.SlideMasterOrder);
            AppendLegacyTextStyleLevels(ownerPart, styles.AddP("titleStyle"), title);
            AppendLegacyTextStyleLevels(ownerPart, styles.AddP("bodyStyle"), body);
            AppendLegacyTextStyleLevels(ownerPart, styles.AddP("otherStyle"), other);
        }

        private static LegacyPptTextMasterStyle? FindUsableStyle(
            IReadOnlyList<LegacyPptTextMasterStyle> styles,
            params LegacyPptTextType[] preferredTypes) {
            foreach (LegacyPptTextType type in preferredTypes) {
                LegacyPptTextMasterStyle? style = styles.FirstOrDefault(candidate =>
                    candidate.TextType == type && !candidate.IsTruncated && candidate.Levels.Count != 0);
                if (style != null) return style;
            }
            return null;
        }

        private static void AppendLegacyTextStyleLevels(PowerPointPartRef ownerPart,
            OfficeXmlNode target,
            LegacyPptTextMasterStyle? source) {
            if (source == null) return;
            foreach (LegacyPptTextMasterStyleLevel level in source.Levels.OrderBy(item => item.Level)) {
                OfficeXmlNode properties = target.AddA(GetLevelParagraphPropertiesName(level.Level));
                LegacyPptTextProjection.ApplyParagraphFormatting(properties,
                    level.ParagraphProperties, includeLevel: false,
                    pictureBullet => ProjectLegacyPictureBullet(ownerPart,
                        pictureBullet));
                LegacyPptTextProjection.AppendDefaultRunProperties(properties, level.CharacterProperties);
            }
        }

        private static string GetLevelParagraphPropertiesName(ushort level) =>
            level switch {
                0 => "lvl1pPr",
                1 => "lvl2pPr",
                2 => "lvl3pPr",
                3 => "lvl4pPr",
                4 => "lvl5pPr",
                _ => throw new ArgumentOutOfRangeException(nameof(level))
            };
    }
}
