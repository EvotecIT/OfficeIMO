using OfficeIMO.PowerPoint.LegacyPpt.Model;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Projects decoded binary text runs into native DrawingML text and fingerprints formatting.</summary>
    internal static class LegacyPptTextProjection {
        internal static void Apply(P.Shape shape, LegacyPptTextBody source,
            Func<LegacyPptInteraction, IReadOnlyList<OpenXmlElement>>? projectInteraction = null) {
            Apply(shape, source, frame: null, projectInteraction,
                projectPictureBullet: null);
        }

        internal static void Apply(P.Shape shape, LegacyPptTextBody source,
            LegacyPptTextFrameProperties? frame,
            Func<LegacyPptInteraction, IReadOnlyList<OpenXmlElement>>?
                projectInteraction = null,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet = null) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (source.HasExplicitCharacterFormatting
                || source.HasParagraphFormatting || source.HasInteractions
                || source.Fields.Count > 0 || source.HasLanguageInformation) {
                shape.TextBody = CreateTextBody(source, frame,
                    projectInteraction, projectPictureBullet);
                return;
            }
            ApplyTextFrame(shape.TextBody?.BodyProperties, frame);
        }

        internal static P.TextBody CreateTextBody(LegacyPptTextBody source) =>
            CreateTextBody(source, frame: null, projectInteraction: null,
                projectPictureBullet: null);

        internal static P.TextBody CreateTextBody(LegacyPptTextBody source,
            Func<LegacyPptInteraction, IReadOnlyList<OpenXmlElement>>? projectInteraction) {
            return CreateTextBody(source, frame: null, projectInteraction,
                projectPictureBullet: null);
        }

        internal static P.TextBody CreateTextBody(LegacyPptTextBody source,
            LegacyPptTextFrameProperties? frame,
            Func<LegacyPptInteraction, IReadOnlyList<OpenXmlElement>>?
                projectInteraction,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet = null) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            var textBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
            ApplyTextFrame(textBody.BodyProperties, frame);
            string[] paragraphs = source.Text.Split(new[] { '\n' }, StringSplitOptions.None);
            int paragraphStart = 0;
            foreach (string paragraphText in paragraphs) {
                int paragraphEnd = checked(paragraphStart + paragraphText.Length);
                var paragraph = new A.Paragraph();
                ApplyParagraphProperties(paragraph, source, paragraphStart,
                    projectPictureBullet);
                AppendParagraphRuns(paragraph, source, paragraphStart, paragraphEnd,
                    projectInteraction);
                if (!paragraph.ChildElements.Any(child => child is A.Run
                        or A.Field or A.Break)) {
                    paragraph.Append(new A.Run(new A.Text(string.Empty)));
                }
                ApplyParagraphEndLanguage(paragraph, source, paragraphEnd);
                textBody.Append(paragraph);
                paragraphStart = checked(paragraphEnd + 1);
            }
            return textBody;
        }

        internal static void ApplyTextFrame(A.BodyProperties? target,
            LegacyPptTextFrameProperties? source) {
            if (target == null || source == null) return;
            if (source.AutoTextMargin != true) {
                if (source.LeftInsetEmus.HasValue) {
                    target.LeftInset = source.LeftInsetEmus.Value;
                }
                if (source.TopInsetEmus.HasValue) {
                    target.TopInset = source.TopInsetEmus.Value;
                }
                if (source.RightInsetEmus.HasValue) {
                    target.RightInset = source.RightInsetEmus.Value;
                }
                if (source.BottomInsetEmus.HasValue) {
                    target.BottomInset = source.BottomInsetEmus.Value;
                }
            }
            if (source.WrapMode.HasValue) {
                target.Wrap = source.WrapMode.Value == 2
                    ? A.TextWrappingValues.None
                    : A.TextWrappingValues.Square;
            }
            if (source.AnchorMode.HasValue) {
                uint anchor = source.AnchorMode.Value;
                target.Anchor = anchor switch {
                    1U or 4U => A.TextAnchoringTypeValues.Center,
                    2U or 5U or 7U or 9U =>
                        A.TextAnchoringTypeValues.Bottom,
                    _ => A.TextAnchoringTypeValues.Top
                };
                if (anchor is 3U or 4U or 5U or 8U or 9U) {
                    target.AnchorCenter = true;
                }
            }
            if (source.TextFlow.HasValue) {
                target.Vertical = source.TextFlow.Value switch {
                    0U or 4U => A.TextVerticalValues.Horizontal,
                    2U => A.TextVerticalValues.Vertical270,
                    _ => A.TextVerticalValues.Vertical
                };
            }
            if (source.FitShapeToText.HasValue) {
                target.Append(source.FitShapeToText.Value
                    ? new A.ShapeAutoFit()
                    : new A.NoAutoFit());
            }
        }

        private static void ApplyParagraphProperties(A.Paragraph paragraph, LegacyPptTextBody source,
            int paragraphStart,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet) {
            LegacyPptParagraphRun? run = source.ParagraphRuns.FirstOrDefault(item =>
                paragraphStart >= item.Start && paragraphStart < item.Start + item.Length);
            if ((run == null || !run.HasExplicitFormatting)
                && source.Ruler?.HasFormatting != true) return;
            ushort level = run?.IndentLevel ?? 0;
            var properties = new A.ParagraphProperties { Level = level };
            ApplyRulerProperties(properties, source.Ruler, level);
            if (run != null) ApplyParagraphFormatting(properties, run,
                includeLevel: true, projectPictureBullet);
            if (run == null || run.TabStops.Count == 0) {
                AppendTabStops(properties, source.Ruler?.TabStops ?? Array.Empty<LegacyPptTabStop>());
            }
            paragraph.Append(properties);
        }

        internal static void ApplyParagraphFormatting(A.TextParagraphPropertiesType properties,
            LegacyPptParagraphRun run, bool includeLevel,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet = null) {
            if (properties == null) throw new ArgumentNullException(nameof(properties));
            if (run == null) throw new ArgumentNullException(nameof(run));
            if (includeLevel) properties.Level = run.IndentLevel;
            if (run.LeftMargin.HasValue) properties.LeftMargin = ToEmus(run.LeftMargin.Value);
            if (run.Indent.HasValue) properties.Indent = ToEmus(run.Indent.Value);
            if (run.DefaultTabSize >= 0) {
                properties.DefaultTabSize = ToEmus(run.DefaultTabSize.Value);
            }
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
            AppendBulletProperties(properties, run, projectPictureBullet);
            AppendTabStops(properties, run.TabStops);
        }

        private static void ApplyRulerProperties(A.TextParagraphPropertiesType properties,
            LegacyPptTextRuler? ruler, ushort level) {
            if (ruler == null) return;
            LegacyPptTextRulerLevel? rulerLevel = ruler.FindLevel(level);
            if (rulerLevel?.LeftMargin != null) {
                properties.LeftMargin = ToEmus(rulerLevel.LeftMargin.Value);
            }
            if (rulerLevel?.Indent != null) properties.Indent = ToEmus(rulerLevel.Indent.Value);
            if (ruler.DefaultTabSize >= 0) {
                properties.DefaultTabSize = ToEmus(ruler.DefaultTabSize.Value);
            }
        }

        private static void AppendBulletProperties(A.TextParagraphPropertiesType properties,
            LegacyPptParagraphRun run,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet) {
            if (run.PictureBullet != null && projectPictureBullet != null) {
                string? relationshipId = projectPictureBullet(
                    run.PictureBullet);
                if (!string.IsNullOrWhiteSpace(relationshipId)) {
                    AppendBulletDecorationProperties(properties, run);
                    properties.Append(new A.PictureBullet(
                        new A.Blip { Embed = relationshipId }));
                    return;
                }
            }
            if (run.HasAutoNumber == true
                && run.AutoNumberScheme.HasValue) {
                AppendBulletDecorationProperties(properties, run);
                properties.Append(new A.AutoNumberedBullet {
                    Type = MapAutoNumberScheme(run.AutoNumberScheme.Value),
                    StartAt = run.AutoNumberStartAt ?? 1
                });
                return;
            }
            if (run.HasBullet == false) {
                properties.Append(new A.NoBullet());
                return;
            }
            AppendBulletDecorationProperties(properties, run);
            if (run.HasBullet == true && run.BulletCharacter.HasValue) {
                properties.Append(new A.CharacterBullet { Char = run.BulletCharacter.Value.ToString() });
            }
        }

        private static void AppendBulletDecorationProperties(
            A.TextParagraphPropertiesType properties,
            LegacyPptParagraphRun run) {
            if (run.BulletHasColor == false) {
                properties.Append(new A.BulletColorText());
            } else if (run.BulletHasColor == true) {
                OpenXmlElement? color = run.BulletColorSchemeIndex.HasValue
                    ? CreateSchemeColor(run.BulletColorSchemeIndex.Value)
                    : run.BulletColor != null
                        ? new A.RgbColorModelHex { Val = run.BulletColor }
                        : null;
                if (color != null) properties.Append(new A.BulletColor(color));
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
        }

        internal static A.TextAutoNumberSchemeValues MapAutoNumberScheme(
            LegacyPptAutoNumberScheme value) => value switch {
                LegacyPptAutoNumberScheme.AlphaLowerPeriod => A.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod,
                LegacyPptAutoNumberScheme.AlphaUpperPeriod => A.TextAutoNumberSchemeValues.AlphaUpperCharacterPeriod,
                LegacyPptAutoNumberScheme.ArabicParenRight => A.TextAutoNumberSchemeValues.ArabicParenR,
                LegacyPptAutoNumberScheme.ArabicPeriod => A.TextAutoNumberSchemeValues.ArabicPeriod,
                LegacyPptAutoNumberScheme.RomanLowerParenBoth => A.TextAutoNumberSchemeValues.RomanLowerCharacterParenBoth,
                LegacyPptAutoNumberScheme.RomanLowerParenRight => A.TextAutoNumberSchemeValues.RomanLowerCharacterParenR,
                LegacyPptAutoNumberScheme.RomanLowerPeriod => A.TextAutoNumberSchemeValues.RomanLowerCharacterPeriod,
                LegacyPptAutoNumberScheme.RomanUpperPeriod => A.TextAutoNumberSchemeValues.RomanUpperCharacterPeriod,
                LegacyPptAutoNumberScheme.AlphaLowerParenBoth => A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenBoth,
                LegacyPptAutoNumberScheme.AlphaLowerParenRight => A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenR,
                LegacyPptAutoNumberScheme.AlphaUpperParenBoth => A.TextAutoNumberSchemeValues.AlphaUpperCharacterParenBoth,
                LegacyPptAutoNumberScheme.AlphaUpperParenRight => A.TextAutoNumberSchemeValues.AlphaUpperCharacterParenR,
                LegacyPptAutoNumberScheme.ArabicParenBoth => A.TextAutoNumberSchemeValues.ArabicParenBoth,
                LegacyPptAutoNumberScheme.ArabicPlain => A.TextAutoNumberSchemeValues.ArabicPlain,
                LegacyPptAutoNumberScheme.RomanUpperParenBoth => A.TextAutoNumberSchemeValues.RomanUpperCharacterParenBoth,
                LegacyPptAutoNumberScheme.RomanUpperParenRight => A.TextAutoNumberSchemeValues.RomanUpperCharacterParenR,
                LegacyPptAutoNumberScheme.SimplifiedChinesePlain => A.TextAutoNumberSchemeValues.EastAsianSimplifiedChinesePlain,
                LegacyPptAutoNumberScheme.SimplifiedChinesePeriod => A.TextAutoNumberSchemeValues.EastAsianSimplifiedChinesePeriod,
                LegacyPptAutoNumberScheme.CircleNumberDoubleBytePlain => A.TextAutoNumberSchemeValues.CircleNumberDoubleBytePlain,
                LegacyPptAutoNumberScheme.CircleNumberWingdingsWhitePlain => A.TextAutoNumberSchemeValues.CircleNumberWingdingsWhitePlain,
                LegacyPptAutoNumberScheme.CircleNumberWingdingsBlackPlain => A.TextAutoNumberSchemeValues.CircleNumberWingdingsBlackPlain,
                LegacyPptAutoNumberScheme.TraditionalChinesePlain => A.TextAutoNumberSchemeValues.EastAsianTraditionalChinesePlain,
                LegacyPptAutoNumberScheme.TraditionalChinesePeriod => A.TextAutoNumberSchemeValues.EastAsianTraditionalChinesePeriod,
                LegacyPptAutoNumberScheme.Arabic1Minus => A.TextAutoNumberSchemeValues.Arabic1Minus,
                LegacyPptAutoNumberScheme.Arabic2Minus => A.TextAutoNumberSchemeValues.Arabic2Minus,
                LegacyPptAutoNumberScheme.Hebrew2Minus => A.TextAutoNumberSchemeValues.Hebrew2Minus,
                LegacyPptAutoNumberScheme.JapaneseKoreanPlain => A.TextAutoNumberSchemeValues.EastAsianJapaneseKoreanPlain,
                LegacyPptAutoNumberScheme.JapaneseKoreanPeriod => A.TextAutoNumberSchemeValues.EastAsianJapaneseKoreanPeriod,
                LegacyPptAutoNumberScheme.ArabicDoubleBytePlain => A.TextAutoNumberSchemeValues.ArabicDoubleBytePlain,
                LegacyPptAutoNumberScheme.ArabicDoubleBytePeriod => A.TextAutoNumberSchemeValues.ArabicDoubleBytePeriod,
                LegacyPptAutoNumberScheme.ThaiAlphaPeriod => A.TextAutoNumberSchemeValues.ThaiAlphaPeriod,
                LegacyPptAutoNumberScheme.ThaiAlphaParenRight => A.TextAutoNumberSchemeValues.ThaiAlphaParenthesisRight,
                LegacyPptAutoNumberScheme.ThaiAlphaParenBoth => A.TextAutoNumberSchemeValues.ThaiAlphaParenthesisBoth,
                LegacyPptAutoNumberScheme.ThaiNumberPeriod => A.TextAutoNumberSchemeValues.ThaiNumberPeriod,
                LegacyPptAutoNumberScheme.ThaiNumberParenRight => A.TextAutoNumberSchemeValues.ThaiNumberParenthesisRight,
                LegacyPptAutoNumberScheme.ThaiNumberParenBoth => A.TextAutoNumberSchemeValues.ThaiNumberParenthesisBoth,
                LegacyPptAutoNumberScheme.HindiAlphaPeriod => A.TextAutoNumberSchemeValues.HindiAlphaPeriod,
                LegacyPptAutoNumberScheme.HindiNumberPeriod => A.TextAutoNumberSchemeValues.HindiNumPeriod,
                LegacyPptAutoNumberScheme.JapaneseDoubleBytePeriod => A.TextAutoNumberSchemeValues.EastAsianJapaneseDoubleBytePeriod,
                LegacyPptAutoNumberScheme.HindiNumberParenRight => A.TextAutoNumberSchemeValues.HindiNumberParenthesisRight,
                LegacyPptAutoNumberScheme.HindiAlpha1Period => A.TextAutoNumberSchemeValues.HindiAlpha1Period,
                _ => throw new ArgumentOutOfRangeException(nameof(value))
            };

        private static void AppendTabStops(A.TextParagraphPropertiesType properties,
            IReadOnlyList<LegacyPptTabStop> tabStops) {
            if (tabStops.Count == 0) return;
            var list = new A.TabStopList();
            foreach (LegacyPptTabStop tabStop in tabStops) {
                list.Append(new A.TabStop {
                    Position = ToEmus(tabStop.Position),
                    Alignment = MapTabAlignment(tabStop.Alignment)
                });
            }
            properties.Append(list);
        }

        private static int ToEmus(short masterUnits) => checked((int)Math.Round(
            masterUnits * 1587.5D, MidpointRounding.AwayFromZero));

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

        private static A.TextTabAlignmentValues MapTabAlignment(LegacyPptTabAlignment value) => value switch {
            LegacyPptTabAlignment.Left => A.TextTabAlignmentValues.Left,
            LegacyPptTabAlignment.Center => A.TextTabAlignmentValues.Center,
            LegacyPptTabAlignment.Right => A.TextTabAlignmentValues.Right,
            LegacyPptTabAlignment.Decimal => A.TextTabAlignmentValues.Decimal,
            _ => throw new ArgumentOutOfRangeException(nameof(value))
        };

        internal static string CreateFormattingFingerprint(
            P.TextBody? textBody, OpenXmlPart? ownerPart = null) {
            if (textBody == null) return string.Empty;
            var clone = (P.TextBody)textBody.CloneNode(true);
            if (clone.BodyProperties != null) {
                clone.ReplaceChild(new A.BodyProperties(),
                    clone.BodyProperties);
            }
            foreach (A.Text text in clone.Descendants<A.Text>()) text.Text = string.Empty;
            foreach (A.RunProperties properties in clone.Descendants<A.RunProperties>()
                         .ToArray()) {
                properties.RemoveAllChildren<A.HyperlinkOnClick>();
                properties.RemoveAllChildren<A.HyperlinkOnMouseOver>();
                if (!properties.HasAttributes && !properties.HasChildren) properties.Remove();
            }
            return clone.OuterXml + CreatePictureBulletImageFingerprint(
                textBody, ownerPart);
        }

        internal static string CreatePictureBulletImageFingerprint(
            OpenXmlElement? root, OpenXmlPart? ownerPart) {
            if (root == null || ownerPart == null) return string.Empty;
            var result = new System.Text.StringBuilder();
            foreach (A.PictureBullet bullet in root
                         .Descendants<A.PictureBullet>()) {
                string? relationshipId = bullet.GetFirstChild<A.Blip>()?
                    .Embed?.Value;
                result.Append("|buBlip:").Append(relationshipId);
                IdPartPair? pair = ownerPart.Parts.FirstOrDefault(item =>
                    string.Equals(item.RelationshipId, relationshipId,
                        StringComparison.Ordinal));
                if (pair?.OpenXmlPart is not ImagePart imagePart) {
                    result.Append(":missing");
                    continue;
                }
                result.Append(':').Append(imagePart.ContentType).Append(':');
                using Stream stream = imagePart.GetStream(FileMode.Open,
                    FileAccess.Read);
                using SHA256 sha = SHA256.Create();
                result.Append(Convert.ToBase64String(
                    sha.ComputeHash(stream)));
            }
            return result.ToString();
        }

        internal static string CreateTextFrameFingerprint(
            P.TextBody? textBody) => textBody?.BodyProperties?.OuterXml
                ?? string.Empty;

        private static void AppendParagraphRuns(A.Paragraph paragraph, LegacyPptTextBody source,
            int paragraphStart, int paragraphEnd,
            Func<LegacyPptInteraction, IReadOnlyList<OpenXmlElement>>? projectInteraction) {
            var boundaries = new SortedSet<int> { paragraphStart, paragraphEnd };
            foreach (LegacyPptCharacterRun run in source.CharacterRuns) {
                AddClippedBoundaries(boundaries, run.Start,
                    checked(run.Start + run.Length), paragraphStart, paragraphEnd);
            }
            foreach (LegacyPptTextLanguageRun run in source.LanguageRuns) {
                AddClippedBoundaries(boundaries, run.Start,
                    checked(run.Start + run.Length), paragraphStart,
                    paragraphEnd);
            }
            foreach (LegacyPptTextInteraction interaction in source.Interactions) {
                AddClippedBoundaries(boundaries, interaction.Start,
                    checked(interaction.Start + interaction.Length), paragraphStart, paragraphEnd);
            }
            foreach (LegacyPptTextField field in source.Fields) {
                AddClippedBoundaries(boundaries, field.Position,
                    checked(field.Position + 1), paragraphStart,
                    paragraphEnd);
            }
            int[] values = boundaries.ToArray();
            for (int index = 0; index < values.Length - 1; index++) {
                int start = values[index];
                int end = values[index + 1];
                if (end <= start) continue;
                LegacyPptCharacterRun? formatting = source.CharacterRuns.FirstOrDefault(run =>
                    start >= run.Start && start < run.Start + run.Length);
                LegacyPptTextLanguageRun? language = source.LanguageRuns
                    .FirstOrDefault(run => start >= run.Start
                        && start < run.Start + run.Length);
                LegacyPptInteraction[] interactions = source.Interactions.Where(item =>
                        start >= item.Start && end <= item.Start + item.Length)
                    .Select(item => item.Interaction)
                    .ToArray();
                LegacyPptTextField? field = end == start + 1
                    ? source.Fields.FirstOrDefault(item =>
                        item.Position == start)
                    : null;
                AppendRun(paragraph,
                    source.Text.Substring(start, end - start), formatting,
                    language,
                    field, interactions, projectInteraction);
            }
        }

        private static void AddClippedBoundaries(ISet<int> boundaries, int start, int end,
            int paragraphStart, int paragraphEnd) {
            int clippedStart = Math.Max(paragraphStart, start);
            int clippedEnd = Math.Min(paragraphEnd, end);
            if (clippedEnd <= clippedStart) return;
            boundaries.Add(clippedStart);
            boundaries.Add(clippedEnd);
        }

        private static void AppendRun(A.Paragraph paragraph, string text,
            LegacyPptCharacterRun? source,
            LegacyPptTextLanguageRun? language,
            LegacyPptTextField? field,
            IReadOnlyList<LegacyPptInteraction> interactions,
            Func<LegacyPptInteraction, IReadOnlyList<OpenXmlElement>>? projectInteraction) {
            A.RunProperties? properties = CreateRunProperties(source,
                language);
            if (projectInteraction != null && interactions.Count > 0) {
                properties ??= new A.RunProperties();
                foreach (LegacyPptInteraction interaction in interactions
                             .GroupBy(item => item.Trigger)
                             .Select(group => group.First())) {
                    foreach (OpenXmlElement element in projectInteraction(interaction)) {
                        properties.Append(element);
                    }
                }
            }
            if (field != null) {
                var nativeField = new A.Field {
                    Id = CreateFieldId(field),
                    Type = MapFieldType(field)
                };
                if (properties != null) {
                    nativeField.RunProperties = (A.RunProperties)properties
                        .CloneNode(true);
                }
                nativeField.Text = new A.Text(text);
                paragraph.Append(nativeField);
                return;
            }
            string[] segments = text.Split(new[] { '\v' },
                StringSplitOptions.None);
            for (int index = 0; index < segments.Length; index++) {
                if (segments[index].Length > 0) {
                    var run = new A.Run(new A.Text(segments[index]));
                    if (properties != null) {
                        run.PrependChild((A.RunProperties)properties
                            .CloneNode(true));
                    }
                    paragraph.Append(run);
                }
                if (index < segments.Length - 1) {
                    var lineBreak = new A.Break();
                    if (properties != null) {
                        lineBreak.RunProperties = (A.RunProperties)properties
                            .CloneNode(true);
                    }
                    paragraph.Append(lineBreak);
                }
            }
        }

        private static string MapFieldType(LegacyPptTextField field) =>
            field.Kind switch {
                LegacyPptTextFieldKind.SlideNumber => "slidenum",
                LegacyPptTextFieldKind.DateTime => "datetime"
                    + checked(field.DateTimeFormatIndex.GetValueOrDefault()
                        + 1).ToString(System.Globalization
                        .CultureInfo.InvariantCulture),
                LegacyPptTextFieldKind.GenericDate =>
                    "datetimeFigureOut",
                LegacyPptTextFieldKind.Header => "header",
                LegacyPptTextFieldKind.Footer => "footer",
                LegacyPptTextFieldKind.RtfDateTime => "datetimeRtf:"
                    + Convert.ToBase64String(System.Text.Encoding.Unicode
                        .GetBytes(field.RtfFormat ?? string.Empty)),
                _ => throw new ArgumentOutOfRangeException(nameof(field))
            };

        private static string CreateFieldId(LegacyPptTextField field) {
            string value = field.Position.ToString(System.Globalization
                    .CultureInfo.InvariantCulture)
                + "|" + MapFieldType(field);
            using SHA256 sha = SHA256.Create();
            byte[] hash = sha.ComputeHash(System.Text.Encoding.UTF8
                .GetBytes(value));
            var guidBytes = new byte[16];
            Buffer.BlockCopy(hash, 0, guidBytes, 0, guidBytes.Length);
            return new Guid(guidBytes).ToString("B").ToUpperInvariant();
        }

        private static A.RunProperties? CreateRunProperties(
            LegacyPptCharacterRun? source,
            LegacyPptTextLanguageRun? language) {
            if ((source == null || !HasNativeCharacterFormatting(source))
                && !HasNativeLanguageInformation(language)) {
                return null;
            }
            var properties = new A.RunProperties();
            if (source != null) ApplyCharacterFormatting(properties, source);
            ApplyLanguageInformation(properties, language);
            return properties;
        }

        internal static A.DefaultRunProperties? CreateDefaultRunProperties(
            LegacyPptCharacterRun source) {
            if (!HasNativeCharacterFormatting(source)) return null;
            var properties = new A.DefaultRunProperties();
            ApplyCharacterFormatting(properties, source);
            return properties;
        }

        private static bool HasNativeCharacterFormatting(LegacyPptCharacterRun source) =>
            source.Bold.HasValue || source.Italic.HasValue
                || source.Underline.HasValue || source.Kumi.HasValue || source.FontSizePoints.HasValue
                || source.Color != null || source.BaselinePositionPercent.HasValue
                || source.Typeface != null || source.AnsiTypeface != null
                || source.OldEastAsianTypeface != null || source.SymbolTypeface != null;

        private static bool HasNativeLanguageInformation(
            LegacyPptTextLanguageRun? source) => source != null
                && (source.Language != null
                    || source.AlternativeLanguage != null || source.NoProof
                    || source.SpellingError.HasValue
                    || source.NeedsRecheck.HasValue);

        private static void ApplyLanguageInformation(
            A.TextCharacterPropertiesType properties,
            LegacyPptTextLanguageRun? source) {
            if (source == null) return;
            if (source.Language != null) properties.Language = source.Language;
            if (source.AlternativeLanguage != null) {
                properties.AlternativeLanguage = source.AlternativeLanguage;
            }
            if (source.NoProof) properties.NoProof = true;
            if (source.SpellingError.HasValue) {
                properties.SpellingError = source.SpellingError.Value;
            }
            if (source.NeedsRecheck.HasValue) {
                properties.Dirty = source.NeedsRecheck.Value;
            }
        }

        private static void ApplyParagraphEndLanguage(A.Paragraph paragraph,
            LegacyPptTextBody source, int paragraphEnd) {
            LegacyPptTextLanguageRun? language = source.LanguageRuns
                .FirstOrDefault(run => paragraphEnd >= run.Start
                    && paragraphEnd < run.Start + run.Length);
            if (!HasNativeLanguageInformation(language)) return;
            var properties = new A.EndParagraphRunProperties();
            ApplyLanguageInformation(properties, language);
            paragraph.Append(properties);
        }

        private static void ApplyCharacterFormatting(A.TextCharacterPropertiesType properties,
            LegacyPptCharacterRun source) {
            if (source.Bold.HasValue) properties.Bold = source.Bold.Value;
            if (source.Italic.HasValue) properties.Italic = source.Italic.Value;
            if (source.Kumi.HasValue) properties.Kumimoji = source.Kumi.Value;
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
            if (source.ColorSchemeIndex.HasValue) {
                properties.Append(new A.SolidFill(
                    CreateSchemeColor(source.ColorSchemeIndex.Value)));
            } else if (source.Color != null) {
                properties.Append(new A.SolidFill(
                    new A.RgbColorModelHex { Val = source.Color }));
            }
            string? latinTypeface = source.AnsiTypeface ?? source.Typeface;
            if (latinTypeface != null) properties.Append(new A.LatinFont { Typeface = latinTypeface });
            if (source.OldEastAsianTypeface != null) {
                properties.Append(new A.EastAsianFont { Typeface = source.OldEastAsianTypeface });
            }
            if (source.SymbolTypeface != null) {
                properties.Append(new A.SymbolFont { Typeface = source.SymbolTypeface });
            }
        }

        private static A.SchemeColor CreateSchemeColor(byte index) => new() {
            Val = index switch {
                0 => A.SchemeColorValues.Background1,
                1 => A.SchemeColorValues.Text1,
                2 => A.SchemeColorValues.Accent4,
                3 => A.SchemeColorValues.Text2,
                4 => A.SchemeColorValues.Background2,
                5 => A.SchemeColorValues.Accent1,
                6 => A.SchemeColorValues.Accent2,
                7 => A.SchemeColorValues.Accent3,
                _ => throw new ArgumentOutOfRangeException(nameof(index))
            }
        };
    }
}
