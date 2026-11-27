using System.Security.Cryptography;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Projects decoded binary text runs into native DrawingML text and fingerprints formatting.</summary>
    internal static partial class LegacyPptTextProjection {
        /// <summary>Adds the markup of an interaction (hyperlink elements) to the run properties it applies to.</summary>
        internal delegate void InteractionProjector(LegacyPptInteraction interaction, OfficeXmlNode runProperties);

        private const int A = LegacyPptXml.A;

        /// <summary>
        /// Gives a projected <c>p:sp</c> the binary text: a new text body when the text carries formatting, interactions,
        /// fields or language runs; otherwise only the text frame on the body the shape has.
        /// </summary>
        internal static void Apply(OfficeXmlNode shape, LegacyPptTextBody source,
            LegacyPptTextFrameProperties? frame,
            InteractionProjector? projectInteraction = null,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet = null) {
            if (shape.IsNone) throw new ArgumentNullException(nameof(shape));
            ArgumentNullException.ThrowIfNull(source);
            if (HasRichText(source)) {
                OfficeXmlNode body = shape.SetChild(LegacyPptXml.P, "txBody", LegacyPptImportXml.ShapeOrder);
                WriteTextBody(body, source, frame, projectInteraction, projectPictureBullet);
                return;
            }
            ApplyTextFrame(shape.Child(LegacyPptXml.P, "txBody").Child(A, "bodyPr"), frame);
        }

        /// <summary>Whether the text needs a text body of its own rather than the plain text the engine writes.</summary>
        internal static bool HasRichText(LegacyPptTextBody source) =>
            source.HasExplicitCharacterFormatting || source.HasParagraphFormatting || source.HasInteractions
            || source.Fields.Count > 0 || source.HasLanguageInformation;

        /// <summary>The projected text body of <paramref name="source"/>, as a <c>p:txBody</c> in a tree of its own.</summary>
        internal static OfficeXmlNode CreateTextBody(LegacyPptTextBody source) {
            ArgumentNullException.ThrowIfNull(source);
            OfficeXmlNode body = LegacyPptImportXml.CreateDetached(LegacyPptXml.P, "txBody");
            WriteTextBody(body, source, frame: null, projectInteraction: null, projectPictureBullet: null);
            return body;
        }

        /// <summary>Fills an empty text body element (<c>p:txBody</c> or <c>a:txBody</c>) with the projected text.</summary>
        internal static void WriteTextBody(OfficeXmlNode textBody, LegacyPptTextBody source,
            LegacyPptTextFrameProperties? frame,
            InteractionProjector? projectInteraction,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet) {
            ArgumentNullException.ThrowIfNull(source);
            OfficeXmlNode bodyProperties = textBody.AppendNew(A, "bodyPr");
            textBody.AppendNew(A, "lstStyle");
            ApplyTextFrame(bodyProperties, frame);
            var projectionPlan = new LegacyPptTextProjectionPlan(source);
            string[] paragraphs = source.Text.Split(new[] { '\n' }, StringSplitOptions.None);
            int paragraphStart = 0;
            foreach (string paragraphText in paragraphs) {
                int paragraphEnd = checked(paragraphStart + paragraphText.Length);
                OfficeXmlNode paragraph = textBody.AppendNew(A, "p");
                ApplyParagraphProperties(paragraph, source, projectionPlan,
                    paragraphStart, projectPictureBullet);
                AppendParagraphRuns(paragraph, source, projectionPlan,
                    paragraphStart, paragraphEnd, projectInteraction);
                if (paragraph.Child(A, "r").IsNone && paragraph.Child(A, "fld").IsNone
                    && paragraph.Child(A, "br").IsNone) {
                    paragraph.AppendNew(A, "r").AppendNew(A, "t").SetInnerText(string.Empty);
                }
                ApplyParagraphEndLanguage(paragraph, projectionPlan, paragraphEnd);
                paragraphStart = checked(paragraphEnd + 1);
            }
        }

        internal static void ApplyTextFrame(OfficeXmlNode target,
            LegacyPptTextFrameProperties? source) {
            if (target.IsNone || source == null) return;
            if (source.AutoTextMargin != true) {
                if (source.LeftInsetEmus.HasValue) target.Set("lIns", source.LeftInsetEmus.Value);
                if (source.TopInsetEmus.HasValue) target.Set("tIns", source.TopInsetEmus.Value);
                if (source.RightInsetEmus.HasValue) target.Set("rIns", source.RightInsetEmus.Value);
                if (source.BottomInsetEmus.HasValue) target.Set("bIns", source.BottomInsetEmus.Value);
            }
            if (source.WrapMode.HasValue) {
                target.Set("wrap", source.WrapMode.Value == 2 ? "none" : "square");
            }
            if (source.AnchorMode.HasValue) {
                uint anchor = source.AnchorMode.Value;
                target.Set("anchor", anchor switch {
                    1U or 4U => "ctr",
                    2U or 5U or 7U or 9U => "b",
                    _ => "t"
                });
                if (anchor is 3U or 4U or 5U or 8U or 9U) target.Set("anchorCtr", true);
            }
            if (source.TextFlow.HasValue) {
                target.Set("vert", source.TextFlow.Value switch {
                    0U or 4U => "horz",
                    2U => "vert270",
                    _ => "vert"
                });
            }
            if (source.FitShapeToText.HasValue) {
                target.AppendNew(A, source.FitShapeToText.Value ? "spAutoFit" : "noAutofit");
            }
        }

        private static void ApplyParagraphProperties(OfficeXmlNode paragraph,
            LegacyPptTextBody source,
            LegacyPptTextProjectionPlan projectionPlan, int paragraphStart,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet) {
            LegacyPptParagraphRun? run = projectionPlan.FindParagraphRun(
                paragraphStart);
            if ((run == null || !run.HasExplicitFormatting)
                && source.Ruler?.HasFormatting != true) return;
            ushort level = run?.IndentLevel ?? 0;
            OfficeXmlNode properties = paragraph.AppendNew(A, "pPr").Set("lvl", level);
            ApplyRulerProperties(properties, source.Ruler, level);
            if (run != null) ApplyParagraphFormatting(properties, run,
                includeLevel: true, projectPictureBullet);
            if (run == null || run.TabStops.Count == 0) {
                AppendTabStops(properties, source.Ruler?.TabStops ?? Array.Empty<LegacyPptTabStop>());
            }
        }

        /// <summary>Writes paragraph formatting into paragraph properties (<c>a:pPr</c> or a list style level).</summary>
        internal static void ApplyParagraphFormatting(OfficeXmlNode properties,
            LegacyPptParagraphRun run, bool includeLevel,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet = null) {
            if (properties.IsNone) throw new ArgumentNullException(nameof(properties));
            ArgumentNullException.ThrowIfNull(run);
            if (includeLevel) properties.Set("lvl", run.IndentLevel);
            if (run.LeftMargin.HasValue) properties.Set("marL", ToEmus(run.LeftMargin.Value));
            if (run.Indent.HasValue) properties.Set("indent", ToEmus(run.Indent.Value));
            if (run.DefaultTabSize >= 0) {
                properties.Set("defTabSz", ToEmus(run.DefaultTabSize.Value));
            }
            if (run.Alignment.HasValue) properties.Set("algn", MapAlignment(run.Alignment.Value));
            if (run.FontAlignment.HasValue) {
                properties.Set("fontAlgn", MapFontAlignment(run.FontAlignment.Value));
            }
            if (run.TextDirection.HasValue) {
                properties.Set("rtl", run.TextDirection.Value == LegacyPptTextDirection.RightToLeft);
            }
            if (run.CharacterWrap.HasValue) properties.Set("eaLnBrk", run.CharacterWrap.Value);
            if (run.WordWrap.HasValue) properties.Set("latinLnBrk", !run.WordWrap.Value);
            if (run.Overflow.HasValue) properties.Set("hangingPunct", run.Overflow.Value);
            if (run.LineSpacing.HasValue) {
                AppendSpacing(properties.AppendNew(A, "lnSpc"), run.LineSpacing.Value);
            }
            if (run.SpaceBefore.HasValue) {
                AppendSpacing(properties.AppendNew(A, "spcBef"), run.SpaceBefore.Value);
            }
            if (run.SpaceAfter.HasValue) {
                AppendSpacing(properties.AppendNew(A, "spcAft"), run.SpaceAfter.Value);
            }
            AppendBulletProperties(properties, run, projectPictureBullet);
            AppendTabStops(properties, run.TabStops);
        }

        private static void ApplyRulerProperties(OfficeXmlNode properties,
            LegacyPptTextRuler? ruler, ushort level) {
            if (ruler == null) return;
            LegacyPptTextRulerLevel? rulerLevel = ruler.FindLevel(level);
            if (rulerLevel?.LeftMargin != null) {
                properties.Set("marL", ToEmus(rulerLevel.LeftMargin.Value));
            }
            if (rulerLevel?.Indent != null) properties.Set("indent", ToEmus(rulerLevel.Indent.Value));
            if (ruler.DefaultTabSize >= 0) {
                properties.Set("defTabSz", ToEmus(ruler.DefaultTabSize.Value));
            }
        }

        private static void AppendBulletProperties(OfficeXmlNode properties,
            LegacyPptParagraphRun run,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet) {
            if (run.PictureBullet != null && projectPictureBullet != null) {
                string? relationshipId = projectPictureBullet(
                    run.PictureBullet);
                if (!string.IsNullOrWhiteSpace(relationshipId)) {
                    AppendBulletDecorationProperties(properties, run);
                    properties.AppendNew(A, "buBlip").AppendNew(A, "blip").SetRelationship("embed", relationshipId);
                    return;
                }
            }
            if (run.HasAutoNumber == true
                && run.AutoNumberScheme.HasValue) {
                AppendBulletDecorationProperties(properties, run);
                properties.AppendNew(A, "buAutoNum")
                    .Set("type", MapAutoNumberScheme(run.AutoNumberScheme.Value).ToOpenXmlToken())
                    .Set("startAt", (long)(run.AutoNumberStartAt ?? 1));
                return;
            }
            if (run.HasBullet == false) {
                properties.AppendNew(A, "buNone");
                return;
            }
            AppendBulletDecorationProperties(properties, run);
            if (run.HasBullet == true && run.BulletCharacter.HasValue
                && LegacyPptXmlText.IsValidAttributeCharacter(
                    run.BulletCharacter.Value)) {
                properties.AppendNew(A, "buChar").Set("char", run.BulletCharacter.Value.ToString());
            }
        }

        private static void AppendBulletDecorationProperties(
            OfficeXmlNode properties,
            LegacyPptParagraphRun run) {
            if (run.BulletHasColor == false) {
                properties.AppendNew(A, "buClrTx");
            } else if (run.BulletHasColor == true) {
                if (run.BulletColorSchemeIndex.HasValue) {
                    AppendSchemeColor(properties.AppendNew(A, "buClr"), run.BulletColorSchemeIndex.Value);
                } else if (run.BulletColor != null) {
                    properties.AppendNew(A, "buClr").AppendNew(A, "srgbClr").Set("val", run.BulletColor);
                }
            }
            if (run.BulletHasSize == false) {
                properties.AppendNew(A, "buSzTx");
            } else if (run.BulletHasSize == true && run.BulletSize.HasValue) {
                short size = run.BulletSize.Value;
                if (size >= 25 && size <= 400) {
                    properties.AppendNew(A, "buSzPct").Set("val", checked(size * 1000));
                } else if (size >= -4000 && size <= -1) {
                    properties.AppendNew(A, "buSzPts").Set("val", checked(-size * 100));
                }
            }
            if (run.BulletHasFont == false) {
                properties.AppendNew(A, "buFontTx");
            } else if (run.BulletHasFont == true) {
                string? typeface = LegacyPptXmlText.SanitizeAttributeValue(
                    run.BulletTypeface);
                if (!string.IsNullOrEmpty(typeface)) {
                    properties.AppendNew(A, "buFont").Set("typeface", typeface);
                }
            }
        }

        internal static PowerPointNumberingScheme MapAutoNumberScheme(
            LegacyPptAutoNumberScheme value) => value switch {
                LegacyPptAutoNumberScheme.AlphaLowerPeriod => PowerPointNumberingScheme.AlphaLowerCharacterPeriod,
                LegacyPptAutoNumberScheme.AlphaUpperPeriod => PowerPointNumberingScheme.AlphaUpperCharacterPeriod,
                LegacyPptAutoNumberScheme.ArabicParenRight => PowerPointNumberingScheme.ArabicParenR,
                LegacyPptAutoNumberScheme.ArabicPeriod => PowerPointNumberingScheme.ArabicPeriod,
                LegacyPptAutoNumberScheme.RomanLowerParenBoth => PowerPointNumberingScheme.RomanLowerCharacterParenBoth,
                LegacyPptAutoNumberScheme.RomanLowerParenRight => PowerPointNumberingScheme.RomanLowerCharacterParenR,
                LegacyPptAutoNumberScheme.RomanLowerPeriod => PowerPointNumberingScheme.RomanLowerCharacterPeriod,
                LegacyPptAutoNumberScheme.RomanUpperPeriod => PowerPointNumberingScheme.RomanUpperCharacterPeriod,
                LegacyPptAutoNumberScheme.AlphaLowerParenBoth => PowerPointNumberingScheme.AlphaLowerCharacterParenBoth,
                LegacyPptAutoNumberScheme.AlphaLowerParenRight => PowerPointNumberingScheme.AlphaLowerCharacterParenR,
                LegacyPptAutoNumberScheme.AlphaUpperParenBoth => PowerPointNumberingScheme.AlphaUpperCharacterParenBoth,
                LegacyPptAutoNumberScheme.AlphaUpperParenRight => PowerPointNumberingScheme.AlphaUpperCharacterParenR,
                LegacyPptAutoNumberScheme.ArabicParenBoth => PowerPointNumberingScheme.ArabicParenBoth,
                LegacyPptAutoNumberScheme.ArabicPlain => PowerPointNumberingScheme.ArabicPlain,
                LegacyPptAutoNumberScheme.RomanUpperParenBoth => PowerPointNumberingScheme.RomanUpperCharacterParenBoth,
                LegacyPptAutoNumberScheme.RomanUpperParenRight => PowerPointNumberingScheme.RomanUpperCharacterParenR,
                LegacyPptAutoNumberScheme.SimplifiedChinesePlain => PowerPointNumberingScheme.EastAsianSimplifiedChinesePlain,
                LegacyPptAutoNumberScheme.SimplifiedChinesePeriod => PowerPointNumberingScheme.EastAsianSimplifiedChinesePeriod,
                LegacyPptAutoNumberScheme.CircleNumberDoubleBytePlain => PowerPointNumberingScheme.CircleNumberDoubleBytePlain,
                LegacyPptAutoNumberScheme.CircleNumberWingdingsWhitePlain => PowerPointNumberingScheme.CircleNumberWingdingsWhitePlain,
                LegacyPptAutoNumberScheme.CircleNumberWingdingsBlackPlain => PowerPointNumberingScheme.CircleNumberWingdingsBlackPlain,
                LegacyPptAutoNumberScheme.TraditionalChinesePlain => PowerPointNumberingScheme.EastAsianTraditionalChinesePlain,
                LegacyPptAutoNumberScheme.TraditionalChinesePeriod => PowerPointNumberingScheme.EastAsianTraditionalChinesePeriod,
                LegacyPptAutoNumberScheme.Arabic1Minus => PowerPointNumberingScheme.Arabic1Minus,
                LegacyPptAutoNumberScheme.Arabic2Minus => PowerPointNumberingScheme.Arabic2Minus,
                LegacyPptAutoNumberScheme.Hebrew2Minus => PowerPointNumberingScheme.Hebrew2Minus,
                LegacyPptAutoNumberScheme.JapaneseKoreanPlain => PowerPointNumberingScheme.EastAsianJapaneseKoreanPlain,
                LegacyPptAutoNumberScheme.JapaneseKoreanPeriod => PowerPointNumberingScheme.EastAsianJapaneseKoreanPeriod,
                LegacyPptAutoNumberScheme.ArabicDoubleBytePlain => PowerPointNumberingScheme.ArabicDoubleBytePlain,
                LegacyPptAutoNumberScheme.ArabicDoubleBytePeriod => PowerPointNumberingScheme.ArabicDoubleBytePeriod,
                LegacyPptAutoNumberScheme.ThaiAlphaPeriod => PowerPointNumberingScheme.ThaiAlphaPeriod,
                LegacyPptAutoNumberScheme.ThaiAlphaParenRight => PowerPointNumberingScheme.ThaiAlphaParenthesisRight,
                LegacyPptAutoNumberScheme.ThaiAlphaParenBoth => PowerPointNumberingScheme.ThaiAlphaParenthesisBoth,
                LegacyPptAutoNumberScheme.ThaiNumberPeriod => PowerPointNumberingScheme.ThaiNumberPeriod,
                LegacyPptAutoNumberScheme.ThaiNumberParenRight => PowerPointNumberingScheme.ThaiNumberParenthesisRight,
                LegacyPptAutoNumberScheme.ThaiNumberParenBoth => PowerPointNumberingScheme.ThaiNumberParenthesisBoth,
                LegacyPptAutoNumberScheme.HindiAlphaPeriod => PowerPointNumberingScheme.HindiAlphaPeriod,
                LegacyPptAutoNumberScheme.HindiNumberPeriod => PowerPointNumberingScheme.HindiNumPeriod,
                LegacyPptAutoNumberScheme.JapaneseDoubleBytePeriod => PowerPointNumberingScheme.EastAsianJapaneseDoubleBytePeriod,
                LegacyPptAutoNumberScheme.HindiNumberParenRight => PowerPointNumberingScheme.HindiNumberParenthesisRight,
                LegacyPptAutoNumberScheme.HindiAlpha1Period => PowerPointNumberingScheme.HindiAlpha1Period,
                _ => throw new ArgumentOutOfRangeException(nameof(value))
            };

        private static void AppendTabStops(OfficeXmlNode properties,
            IReadOnlyList<LegacyPptTabStop> tabStops) {
            if (tabStops.Count == 0) return;
            OfficeXmlNode list = properties.AppendNew(A, "tabLst");
            foreach (LegacyPptTabStop tabStop in tabStops.Take(PowerPointParagraph.MaximumTabStops)) {
                list.AppendNew(A, "tab")
                    .Set("pos", ToEmus(tabStop.Position))
                    .Set("algn", MapTabAlignment(tabStop.Alignment));
            }
        }

        private static int ToEmus(short masterUnits) => checked((int)Math.Round(
            masterUnits * 1587.5D, MidpointRounding.AwayFromZero));

        private static void AppendSpacing(OfficeXmlNode spacing, short value) {
            if (value >= 0) {
                spacing.AppendNew(A, "spcPct").Set("val", checked(value * 1000));
                return;
            }
            int points = checked((int)Math.Round(-(long)value * 12.5D,
                MidpointRounding.AwayFromZero));
            spacing.AppendNew(A, "spcPts").Set("val", points);
        }

        private static string MapAlignment(LegacyPptTextAlignment value) => value switch {
            LegacyPptTextAlignment.Left => "l",
            LegacyPptTextAlignment.Center => "ctr",
            LegacyPptTextAlignment.Right => "r",
            LegacyPptTextAlignment.Justify => "just",
            LegacyPptTextAlignment.Distributed => "dist",
            LegacyPptTextAlignment.ThaiDistributed => "thaiDist",
            LegacyPptTextAlignment.JustifyLow => "justLow",
            _ => throw new ArgumentOutOfRangeException(nameof(value))
        };

        private static string MapFontAlignment(LegacyPptFontAlignment value) => value switch {
            LegacyPptFontAlignment.Baseline => "base",
            LegacyPptFontAlignment.Hanging => "t",
            LegacyPptFontAlignment.Center => "ctr",
            LegacyPptFontAlignment.Bottom => "b",
            _ => throw new ArgumentOutOfRangeException(nameof(value))
        };

        private static string MapTabAlignment(LegacyPptTabAlignment value) => value switch {
            LegacyPptTabAlignment.Left => "l",
            LegacyPptTabAlignment.Center => "ctr",
            LegacyPptTabAlignment.Right => "r",
            LegacyPptTabAlignment.Decimal => "dec",
            _ => throw new ArgumentOutOfRangeException(nameof(value))
        };

        /// <summary>
        /// A fingerprint of a text body's formatting: its markup with the text frame, the text and run interactions left
        /// out, plus the content of each picture bullet image. Read from a package view; values compare only with other
        /// values of this method.
        /// </summary>
        internal static string CreateFormattingFingerprint(OfficeXmlNode textBody, LegacyPptPartView ownerPart) {
            if (textBody.IsNone) return string.Empty;
            OfficeXmlNode clone = textBody.Clone();
            OfficeXmlNode bodyProperties = clone.Child(A, "bodyPr");
            if (!bodyProperties.IsNone) {
                clone.InsertBefore(clone.CreateElement(A, "bodyPr"), bodyProperties);
                bodyProperties.Remove();
            }
            foreach (OfficeXmlNode text in clone.Descendants(A, "t")) text.SetInnerText(string.Empty);
            foreach (OfficeXmlNode properties in clone.Descendants(A, "rPr")) {
                properties.RemoveChildren(A, "hlinkClick");
                properties.RemoveChildren(A, "hlinkMouseOver");
                if (properties.FirstChild.IsNone && !HasAttributes(properties)) properties.Remove();
            }
            return LegacyPptXml.Digest(clone) + CreatePictureBulletImageFingerprint(textBody, ownerPart);
        }

        /// <summary>The content type and content digest of each picture bullet image below <paramref name="root"/>.</summary>
        internal static string CreatePictureBulletImageFingerprint(OfficeXmlNode root, LegacyPptPartView ownerPart) {
            if (root.IsNone || ownerPart.IsNone) return string.Empty;
            var result = new System.Text.StringBuilder();
            foreach (OfficeXmlNode bullet in root.Descendants(A, "buBlip")) {
                string? relationshipId = bullet.Child(A, "blip").RelationshipId("embed");
                result.Append("|buBlip:").Append(relationshipId);
                LegacyPptPartView image = ownerPart.PartById(relationshipId);
                if (image.IsNone || image.ContentType?.StartsWith("image/", StringComparison.OrdinalIgnoreCase) != true) {
                    result.Append(":missing");
                    continue;
                }
                result.Append(':').Append(image.ContentType).Append(':');
                using SHA256 sha = SHA256.Create();
                result.Append(Convert.ToBase64String(sha.ComputeHash(image.GetBytes())));
            }
            return result.ToString();
        }

        /// <summary>A fingerprint of a text body's frame properties.</summary>
        internal static string CreateTextFrameFingerprint(OfficeXmlNode textBody) {
            OfficeXmlNode bodyProperties = textBody.Child(A, "bodyPr");
            return bodyProperties.IsNone ? string.Empty : LegacyPptXml.Digest(bodyProperties);
        }

        // Whether an element has attributes other than namespace declarations.
        private static bool HasAttributes(OfficeXmlNode element) {
            if (element.Tree is not { } tree) throw new InvalidOperationException("Fingerprints are read from Open XML layer trees.");
            int count = tree.AttributeCountOf(element.Node);
            for (int index = 0; index < count; index++) {
                tree.GetAttributeAt(element.Node, index, out int namespaceId, out string localName, out _);
                if (namespaceId == OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Xmlns
                    || namespaceId == OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.None && localName == "xmlns") continue;
                return true;
            }
            return false;
        }

        private static void AppendParagraphRuns(OfficeXmlNode paragraph,
            LegacyPptTextBody source,
            LegacyPptTextProjectionPlan projectionPlan,
            int paragraphStart, int paragraphEnd,
            InteractionProjector? projectInteraction) {
            IReadOnlyList<int> boundaries = projectionPlan
                .GetParagraphBoundaries(paragraphStart, paragraphEnd);
            for (int index = 0; index < boundaries.Count - 1; index++) {
                int start = boundaries[index];
                int end = boundaries[index + 1];
                if (end <= start) continue;
                LegacyPptCharacterRun? formatting = projectionPlan
                    .FindCharacterRun(start);
                LegacyPptTextLanguageRun? language = projectionPlan
                    .FindLanguageRun(start);
                IReadOnlyList<LegacyPptInteraction> interactions =
                    projectionPlan.GetInteractions(start);
                LegacyPptTextField? field = projectionPlan.FindField(start,
                    end);
                AppendRun(paragraph,
                    source.Text.Substring(start, end - start), formatting,
                    language,
                    field, interactions, projectInteraction);
            }
        }

        private static void AppendRun(OfficeXmlNode paragraph, string text,
            LegacyPptCharacterRun? source,
            LegacyPptTextLanguageRun? language,
            LegacyPptTextField? field,
            IReadOnlyList<LegacyPptInteraction> interactions,
            InteractionProjector? projectInteraction) {
            OfficeXmlNode properties = CreateRunProperties(paragraph, source,
                language);
            if (projectInteraction != null && interactions.Count > 0) {
                if (properties.IsNone) properties = paragraph.CreateElement(A, "rPr");
                foreach (LegacyPptInteraction interaction in interactions
                             .GroupBy(item => item.Trigger)
                             .Select(group => group.First())) {
                    projectInteraction(interaction, properties);
                }
            }
            if (field != null) {
                OfficeXmlNode nativeField = paragraph.AppendNew(A, "fld")
                    .Set("id", CreateFieldId(field))
                    .Set("type", MapFieldType(field));
                if (!properties.IsNone) nativeField.Append(properties.Clone());
                nativeField.AppendNew(A, "t").SetInnerText(text);
                return;
            }
            string[] segments = text.Split(new[] { '\v' },
                StringSplitOptions.None);
            for (int index = 0; index < segments.Length; index++) {
                if (segments[index].Length > 0) {
                    OfficeXmlNode run = paragraph.AppendNew(A, "r");
                    if (!properties.IsNone) run.Append(properties.Clone());
                    run.AppendNew(A, "t").SetInnerText(segments[index]);
                }
                if (index < segments.Length - 1) {
                    OfficeXmlNode lineBreak = paragraph.AppendNew(A, "br");
                    if (!properties.IsNone) lineBreak.Append(properties.Clone());
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

        // Detached run properties created in the paragraph's part, or none when the run has nothing to carry.
        private static OfficeXmlNode CreateRunProperties(OfficeXmlNode context,
            LegacyPptCharacterRun? source,
            LegacyPptTextLanguageRun? language) {
            if ((source == null || !HasNativeCharacterFormatting(source))
                && !HasNativeLanguageInformation(language)) {
                return default;
            }
            OfficeXmlNode properties = context.CreateElement(A, "rPr");
            if (source != null) ApplyCharacterFormatting(properties, source);
            ApplyLanguageInformation(properties, language);
            return properties;
        }

        /// <summary>Appends <c>a:defRPr</c> with the character formatting; false, adding nothing, when there is none.</summary>
        internal static bool AppendDefaultRunProperties(OfficeXmlNode parent,
            LegacyPptCharacterRun source) {
            if (!HasNativeCharacterFormatting(source)) return false;
            ApplyCharacterFormatting(parent.AppendNew(A, "defRPr"), source);
            return true;
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
            OfficeXmlNode properties,
            LegacyPptTextLanguageRun? source) {
            if (source == null) return;
            if (source.Language != null) properties.Set("lang", source.Language);
            if (source.AlternativeLanguage != null) {
                properties.Set("altLang", source.AlternativeLanguage);
            }
            if (source.NoProof) properties.Set("noProof", true);
            if (source.SpellingError.HasValue) {
                properties.Set("err", source.SpellingError.Value);
            }
            if (source.NeedsRecheck.HasValue) {
                properties.Set("dirty", source.NeedsRecheck.Value);
            }
        }

        private static void ApplyParagraphEndLanguage(OfficeXmlNode paragraph,
            LegacyPptTextProjectionPlan projectionPlan,
            int paragraphEnd) {
            LegacyPptTextLanguageRun? language = projectionPlan
                .FindLanguageRun(paragraphEnd);
            if (!HasNativeLanguageInformation(language)) return;
            ApplyLanguageInformation(paragraph.AppendNew(A, "endParaRPr"), language);
        }

        private static void ApplyCharacterFormatting(OfficeXmlNode properties,
            LegacyPptCharacterRun source) {
            if (source.Bold.HasValue) properties.Set("b", source.Bold.Value);
            if (source.Italic.HasValue) properties.Set("i", source.Italic.Value);
            if (source.Kumi.HasValue) properties.Set("kumimoji", source.Kumi.Value);
            if (source.Underline.HasValue) {
                properties.Set("u", source.Underline.Value ? "sng" : "none");
            }
            if (source.FontSizePoints.HasValue && source.FontSizePoints.Value > 0) {
                properties.Set("sz", checked(source.FontSizePoints.Value * 100));
            }
            if (source.BaselinePositionPercent.HasValue) {
                properties.Set("baseline", checked(source.BaselinePositionPercent.Value * 1000));
            }
            if (source.ColorSchemeIndex.HasValue) {
                AppendSchemeColor(properties.AppendNew(A, "solidFill"), source.ColorSchemeIndex.Value);
            } else if (source.Color != null) {
                properties.AppendNew(A, "solidFill").AppendNew(A, "srgbClr").Set("val", source.Color);
            }
            string? latinTypeface = source.AnsiTypeface ?? source.Typeface;
            latinTypeface = LegacyPptXmlText.SanitizeAttributeValue(latinTypeface);
            if (!string.IsNullOrEmpty(latinTypeface)) {
                properties.AppendNew(A, "latin").Set("typeface", latinTypeface);
            }
            string? eastAsianTypeface = LegacyPptXmlText.SanitizeAttributeValue(
                source.OldEastAsianTypeface);
            if (!string.IsNullOrEmpty(eastAsianTypeface)) {
                properties.AppendNew(A, "ea").Set("typeface", eastAsianTypeface);
            }
            string? symbolTypeface = LegacyPptXmlText.SanitizeAttributeValue(
                source.SymbolTypeface);
            if (!string.IsNullOrEmpty(symbolTypeface)) {
                properties.AppendNew(A, "sym").Set("typeface", symbolTypeface);
            }
        }

        private static void AppendSchemeColor(OfficeXmlNode parent, byte index) {
            string value = index switch {
                0 => "bg1",
                1 => "tx1",
                2 => "accent4",
                3 => "tx2",
                4 => "bg2",
                5 => "accent1",
                6 => "accent2",
                7 => "accent3",
                _ => throw new ArgumentOutOfRangeException(nameof(index))
            };
            parent.AppendNew(A, "schemeClr").Set("val", value);
        }
    }
}
