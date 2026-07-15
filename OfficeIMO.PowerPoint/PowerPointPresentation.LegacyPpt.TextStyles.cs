using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyMasterTextStyles(SlideMaster slideMaster,
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
                CreateTitleStyle(title), CreateBodyStyle(body), CreateOtherStyle(other));
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

        private static TitleStyle CreateTitleStyle(LegacyPptTextMasterStyle? source) {
            var target = new TitleStyle();
            AppendLegacyTextStyleLevels(target, source);
            return target;
        }

        private static BodyStyle CreateBodyStyle(LegacyPptTextMasterStyle? source) {
            var target = new BodyStyle();
            AppendLegacyTextStyleLevels(target, source);
            return target;
        }

        private static OtherStyle CreateOtherStyle(LegacyPptTextMasterStyle? source) {
            var target = new OtherStyle();
            AppendLegacyTextStyleLevels(target, source);
            return target;
        }

        private static void AppendLegacyTextStyleLevels(OpenXmlCompositeElement target,
            LegacyPptTextMasterStyle? source) {
            if (source == null) return;
            foreach (LegacyPptTextMasterStyleLevel level in source.Levels.OrderBy(item => item.Level)) {
                A.TextParagraphPropertiesType properties = CreateLevelParagraphProperties(level.Level);
                LegacyPptTextProjection.ApplyParagraphFormatting(properties,
                    level.ParagraphProperties, includeLevel: false);
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
