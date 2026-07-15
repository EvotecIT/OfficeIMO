using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Projects decoded binary text runs into native DrawingML text and fingerprints formatting.</summary>
    internal static class LegacyPptTextProjection {
        internal static void Apply(P.Shape shape, LegacyPptTextBody source) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!source.HasExplicitCharacterFormatting && !source.HasParagraphFormatting) return;
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
                ApplyParagraphProperties(paragraph, source, paragraphStart);
                AppendParagraphRuns(paragraph, source, paragraphStart, paragraphEnd);
                if (!paragraph.Elements<A.Run>().Any()) {
                    paragraph.Append(new A.Run(new A.Text(string.Empty)));
                }
                textBody.Append(paragraph);
                paragraphStart = checked(paragraphEnd + 1);
            }
            return textBody;
        }

        private static void ApplyParagraphProperties(A.Paragraph paragraph, LegacyPptTextBody source,
            int paragraphStart) {
            LegacyPptParagraphRun? run = source.ParagraphRuns.FirstOrDefault(item =>
                paragraphStart >= item.Start && paragraphStart < item.Start + item.Length);
            if (run == null || !run.HasExplicitFormatting) return;
            var properties = new A.ParagraphProperties { Level = run.IndentLevel };
            if (run.Alignment.HasValue) properties.Alignment = MapAlignment(run.Alignment.Value);
            if (run.FontAlignment.HasValue) {
                properties.FontAlignment = MapFontAlignment(run.FontAlignment.Value);
            }
            if (run.TextDirection.HasValue) {
                properties.RightToLeft = run.TextDirection.Value == LegacyPptTextDirection.RightToLeft;
            }
            if (run.CharacterWrap.HasValue) properties.EastAsianLineBreak = run.CharacterWrap.Value;
            if (run.WordWrap.HasValue) properties.LatinLineBreak = !run.WordWrap.Value;
            if (run.Overflow.HasValue) properties.Height = run.Overflow.Value;
            if (run.LineSpacing.HasValue) {
                properties.Append(new A.LineSpacing(CreateSpacing(run.LineSpacing.Value)));
            }
            if (run.SpaceBefore.HasValue) {
                properties.Append(new A.SpaceBefore(CreateSpacing(run.SpaceBefore.Value)));
            }
            if (run.SpaceAfter.HasValue) {
                properties.Append(new A.SpaceAfter(CreateSpacing(run.SpaceAfter.Value)));
            }
            AppendBulletProperties(properties, run);
            paragraph.Append(properties);
        }

        private static void AppendBulletProperties(A.ParagraphProperties properties,
            LegacyPptParagraphRun run) {
            if (run.HasBullet == false) {
                properties.Append(new A.NoBullet());
                return;
            }
            if (run.BulletHasColor == false) {
                properties.Append(new A.BulletColorText());
            } else if (run.BulletHasColor == true && run.BulletColor != null) {
                properties.Append(new A.BulletColor(
                    new A.RgbColorModelHex { Val = run.BulletColor }));
            }
            if (run.BulletHasSize == false) {
                properties.Append(new A.BulletSizeText());
            } else if (run.BulletHasSize == true && run.BulletSize.HasValue) {
                short size = run.BulletSize.Value;
                if (size >= 25 && size <= 400) {
                    properties.Append(new A.BulletSizePercentage { Val = checked(size * 1000) });
                } else if (size >= -4000 && size <= -1) {
                    properties.Append(new A.BulletSizePoints { Val = checked(-size * 100) });
                }
            }
            if (run.BulletHasFont == false) {
                properties.Append(new A.BulletFontText());
            } else if (run.BulletHasFont == true && run.BulletTypeface != null) {
                properties.Append(new A.BulletFont { Typeface = run.BulletTypeface });
            }
            if (run.HasBullet == true && run.BulletCharacter.HasValue) {
                properties.Append(new A.CharacterBullet { Char = run.BulletCharacter.Value.ToString() });
            }
        }

        private static OpenXmlElement CreateSpacing(short value) {
            if (value >= 0) return new A.SpacingPercent { Val = checked(value * 1000) };
            int points = checked((int)Math.Round(-(long)value * 12.5D,
                MidpointRounding.AwayFromZero));
            return new A.SpacingPoints { Val = points };
        }

        private static A.TextAlignmentTypeValues MapAlignment(LegacyPptTextAlignment value) => value switch {
            LegacyPptTextAlignment.Left => A.TextAlignmentTypeValues.Left,
            LegacyPptTextAlignment.Center => A.TextAlignmentTypeValues.Center,
            LegacyPptTextAlignment.Right => A.TextAlignmentTypeValues.Right,
            LegacyPptTextAlignment.Justify => A.TextAlignmentTypeValues.Justified,
            LegacyPptTextAlignment.Distributed => A.TextAlignmentTypeValues.Distributed,
            LegacyPptTextAlignment.ThaiDistributed => A.TextAlignmentTypeValues.ThaiDistributed,
            LegacyPptTextAlignment.JustifyLow => A.TextAlignmentTypeValues.JustifiedLow,
            _ => throw new ArgumentOutOfRangeException(nameof(value))
        };

        private static A.TextFontAlignmentValues MapFontAlignment(LegacyPptFontAlignment value) => value switch {
            LegacyPptFontAlignment.Baseline => A.TextFontAlignmentValues.Baseline,
            LegacyPptFontAlignment.Hanging => A.TextFontAlignmentValues.Top,
            LegacyPptFontAlignment.Center => A.TextFontAlignmentValues.Center,
            LegacyPptFontAlignment.Bottom => A.TextFontAlignmentValues.Bottom,
            _ => throw new ArgumentOutOfRangeException(nameof(value))
        };

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
                || source.Color != null || source.BaselinePositionPercent.HasValue
                || source.Typeface != null || source.AnsiTypeface != null
                || source.OldEastAsianTypeface != null || source.SymbolTypeface != null;
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
            string? latinTypeface = source.AnsiTypeface ?? source.Typeface;
            if (latinTypeface != null) properties.Append(new A.LatinFont { Typeface = latinTypeface });
            if (source.OldEastAsianTypeface != null) {
                properties.Append(new A.EastAsianFont { Typeface = source.OldEastAsianTypeface });
            }
            if (source.SymbolTypeface != null) {
                properties.Append(new A.SymbolFont { Typeface = source.SymbolTypeface });
            }
            return properties;
        }
    }
}
