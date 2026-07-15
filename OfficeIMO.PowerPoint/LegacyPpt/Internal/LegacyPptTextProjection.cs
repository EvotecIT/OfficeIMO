using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Projects decoded binary text runs into native DrawingML text and fingerprints formatting.</summary>
    internal static class LegacyPptTextProjection {
        internal static void Apply(P.Shape shape, LegacyPptTextBody source) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!source.HasExplicitCharacterFormatting) return;
            shape.TextBody = CreateTextBody(source);
        }

        internal static P.TextBody CreateTextBody(LegacyPptTextBody source) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            var textBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
            string[] paragraphs = source.Text.Split(new[] { '\n' }, StringSplitOptions.None);
            int paragraphStart = 0;
            foreach (string paragraphText in paragraphs) {
                int paragraphEnd = checked(paragraphStart + paragraphText.Length);
                var paragraph = new A.Paragraph();
                AppendParagraphRuns(paragraph, source, paragraphStart, paragraphEnd);
                if (!paragraph.Elements<A.Run>().Any()) {
                    paragraph.Append(new A.Run(new A.Text(string.Empty)));
                }
                textBody.Append(paragraph);
                paragraphStart = checked(paragraphEnd + 1);
            }
            return textBody;
        }

        internal static string CreateFormattingFingerprint(P.TextBody? textBody) {
            if (textBody == null) return string.Empty;
            var clone = (P.TextBody)textBody.CloneNode(true);
            foreach (A.Text text in clone.Descendants<A.Text>()) text.Text = string.Empty;
            return clone.OuterXml;
        }

        private static void AppendParagraphRuns(A.Paragraph paragraph, LegacyPptTextBody source,
            int paragraphStart, int paragraphEnd) {
            int cursor = paragraphStart;
            foreach (LegacyPptCharacterRun run in source.CharacterRuns.OrderBy(item => item.Start)) {
                int runEnd = checked(run.Start + run.Length);
                int start = Math.Max(paragraphStart, run.Start);
                int end = Math.Min(paragraphEnd, runEnd);
                if (end <= start) continue;
                if (start > cursor) AppendRun(paragraph, source.Text.Substring(cursor, start - cursor), null);
                AppendRun(paragraph, source.Text.Substring(start, end - start), run);
                cursor = Math.Max(cursor, end);
            }
            if (cursor < paragraphEnd) {
                AppendRun(paragraph, source.Text.Substring(cursor, paragraphEnd - cursor), null);
            }
        }

        private static void AppendRun(A.Paragraph paragraph, string text, LegacyPptCharacterRun? source) {
            var run = new A.Run(new A.Text(text));
            A.RunProperties? properties = source == null ? null : CreateRunProperties(source);
            if (properties != null) run.PrependChild(properties);
            paragraph.Append(run);
        }

        private static A.RunProperties? CreateRunProperties(LegacyPptCharacterRun source) {
            bool hasNativeFormatting = source.Bold.HasValue || source.Italic.HasValue
                || source.Underline.HasValue || source.FontSizePoints.HasValue
                || source.Color != null || source.BaselinePositionPercent.HasValue;
            if (!hasNativeFormatting) return null;
            var properties = new A.RunProperties();
            if (source.Bold.HasValue) properties.Bold = source.Bold.Value;
            if (source.Italic.HasValue) properties.Italic = source.Italic.Value;
            if (source.Underline.HasValue) {
                properties.Underline = source.Underline.Value
                    ? A.TextUnderlineValues.Single
                    : A.TextUnderlineValues.None;
            }
            if (source.FontSizePoints.HasValue && source.FontSizePoints.Value > 0) {
                properties.FontSize = checked(source.FontSizePoints.Value * 100);
            }
            if (source.BaselinePositionPercent.HasValue) {
                properties.Baseline = checked(source.BaselinePositionPercent.Value * 1000);
            }
            if (source.Color != null) {
                properties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = source.Color }));
            }
            return properties;
        }
    }
}
