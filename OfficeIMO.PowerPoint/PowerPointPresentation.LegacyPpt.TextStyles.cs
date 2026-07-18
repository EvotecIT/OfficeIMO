using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyMasterTextStyles(
            SlideMasterPart ownerPart, SlideMaster slideMaster,
            IReadOnlyList<LegacyPptTextMasterStyle> source) {
            if (source.Count == 0) return;
            LegacyPptTextMasterStyle? title = FindUsableStyle(source,
                LegacyPptTextType.Title, LegacyPptTextType.CenterTitle);
            LegacyPptTextMasterStyle? body = FindUsableStyle(source,
                LegacyPptTextType.Body, LegacyPptTextType.CenterBody,
                LegacyPptTextType.HalfBody, LegacyPptTextType.QuarterBody);
            LegacyPptTextMasterStyle? other = FindUsableStyle(source,
                LegacyPptTextType.Other, LegacyPptTextType.Notes);

            slideMaster.TextStyles = new TextStyles(
                CreateTitleStyle(ownerPart, title),
                CreateBodyStyle(ownerPart, body),
                CreateOtherStyle(ownerPart, other));
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

        private static TitleStyle CreateTitleStyle(OpenXmlPart ownerPart,
            LegacyPptTextMasterStyle? source) {
            var target = new TitleStyle();
            AppendLegacyTextStyleLevels(ownerPart, target, source);
            return target;
        }

        private static BodyStyle CreateBodyStyle(OpenXmlPart ownerPart,
            LegacyPptTextMasterStyle? source) {
            var target = new BodyStyle();
            AppendLegacyTextStyleLevels(ownerPart, target, source);
            return target;
        }

        private static OtherStyle CreateOtherStyle(OpenXmlPart ownerPart,
            LegacyPptTextMasterStyle? source) {
            var target = new OtherStyle();
            AppendLegacyTextStyleLevels(ownerPart, target, source);
            return target;
        }

        private static void AppendLegacyTextStyleLevels(OpenXmlPart ownerPart,
            OpenXmlCompositeElement target,
            LegacyPptTextMasterStyle? source) {
            if (source == null) return;
            foreach (LegacyPptTextMasterStyleLevel level in source.Levels.OrderBy(item => item.Level)) {
                A.TextParagraphPropertiesType properties = CreateLevelParagraphProperties(level.Level);
                LegacyPptTextProjection.ApplyParagraphFormatting(properties,
                    level.ParagraphProperties, includeLevel: false,
                    pictureBullet => ProjectLegacyPictureBullet(ownerPart,
                        pictureBullet));
                A.DefaultRunProperties? runProperties = LegacyPptTextProjection
                    .CreateDefaultRunProperties(level.CharacterProperties);
                if (runProperties != null) properties.Append(runProperties);
                target.Append(properties);
            }
        }

        private static A.TextParagraphPropertiesType CreateLevelParagraphProperties(ushort level) =>
            level switch {
                0 => new A.Level1ParagraphProperties(),
                1 => new A.Level2ParagraphProperties(),
                2 => new A.Level3ParagraphProperties(),
                3 => new A.Level4ParagraphProperties(),
                4 => new A.Level5ParagraphProperties(),
                _ => throw new ArgumentOutOfRangeException(nameof(level))
            };
    }
}
