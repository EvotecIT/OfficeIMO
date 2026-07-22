using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf.Writing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private const int MaximumWordListLevel = 8;

    private static void CopyWordStylesAndNumbering(WordDocument source, RtfDocument destination) {
        Styles? styles = source._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles != null) {
            var ids = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int nextId = 0;
            foreach (Style style in styles.Elements<Style>().Where(style => style.StyleId?.Value != null)) {
                ids[style.StyleId!.Value!] = nextId++;
            }

            foreach (Style style in styles.Elements<Style>().Where(style => style.StyleId?.Value != null)) {
                string styleId = style.StyleId!.Value!;
                RtfStyleKind kind = ToRtfStyleKind(style.Type?.Value);
                RtfStyle target = destination.AddStyle(ids[styleId], styleId, kind);
                target.BasedOnStyleId = ResolveStyleReference(style.BasedOn?.Val?.Value, ids);
                target.NextStyleId = ResolveStyleReference(style.NextParagraphStyle?.Val?.Value, ids);
                target.LinkedStyleId = ResolveStyleReference(style.LinkedStyle?.Val?.Value, ids);
                target.AutoUpdate = style.AutoRedefine != null;
                target.Hidden = style.GetFirstChild<StyleHidden>() != null;
                target.Locked = style.Locked != null;
                target.SemiHidden = style.SemiHidden != null;
                target.UnhideWhenUsed = style.UnhideWhenUsed != null;
                target.QuickFormat = style.PrimaryStyle != null;
                target.Priority = style.UIPriority?.Val?.Value;
                CopyWordStyleFormatting(style, target, destination);
            }
        }

        CopyWordNumbering(source, destination);
    }

    private static void CopyWordStyleFormatting(Style source, RtfStyle destination, RtfDocument document) {
        StyleRunProperties? run = source.StyleRunProperties;
        if (run != null) {
            destination.Bold = ReadToggle(run.Bold);
            destination.Italic = ReadToggle(run.Italic);
            if (run.Underline?.Val?.Value is UnderlineValues underline) destination.UnderlineStyle = ToRtfUnderlineStyle(underline);
            if (double.TryParse(run.FontSize?.Val?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints)) destination.FontSize = halfPoints / 2d;
            string? font = run.RunFonts?.Ascii?.Value ?? run.RunFonts?.HighAnsi?.Value ?? run.RunFonts?.EastAsia?.Value;
            if (!string.IsNullOrWhiteSpace(font)) destination.FontId = document.AddFont(font!);
        }

        StyleParagraphProperties? paragraph = source.StyleParagraphProperties;
        if (paragraph == null) return;
        destination.ParagraphAlignment = ToRtfTextAlignment(paragraph.Justification?.Val?.Value);
        destination.LeftIndentTwips = ParseInt(paragraph.Indentation?.Left?.Value);
        destination.RightIndentTwips = ParseInt(paragraph.Indentation?.Right?.Value);
        destination.FirstLineIndentTwips = ParseInt(paragraph.Indentation?.FirstLine?.Value) ?? Negate(ParseInt(paragraph.Indentation?.Hanging?.Value));
        destination.SpaceBeforeTwips = ParseInt(paragraph.SpacingBetweenLines?.Before?.Value);
        destination.SpaceAfterTwips = ParseInt(paragraph.SpacingBetweenLines?.After?.Value);
        destination.LineSpacingTwips = ParseInt(paragraph.SpacingBetweenLines?.Line?.Value);
        destination.PageBreakBefore = ReadToggle(paragraph.PageBreakBefore);
        destination.KeepWithNext = ReadToggle(paragraph.KeepNext);
        destination.KeepLinesTogether = ReadToggle(paragraph.KeepLines);
        destination.OutlineLevel = paragraph.OutlineLevel?.Val?.Value;
    }

    private static void CopyWordNumbering(WordDocument source, RtfDocument destination) {
        Numbering? numbering = source._wordprocessingDocument.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return;
        foreach (AbstractNum abstractNum in numbering.Elements<AbstractNum>()) {
            int? id = abstractNum.AbstractNumberId?.Value;
            if (!id.HasValue) continue;
            RtfListDefinition definition = destination.AddListDefinition(id.Value);
            var seenLevels = new HashSet<int>();
            foreach (Level wordLevel in abstractNum.Elements<Level>().OrderBy(level => level.LevelIndex?.Value ?? 0)) {
                NumberFormatValues? format = wordLevel.NumberingFormat?.Val?.Value;
                int levelIndex = wordLevel.LevelIndex?.Value ?? definition.Levels.Count;
                ValidateWordListLevel(levelIndex);
                if (!seenLevels.Add(levelIndex)) {
                    throw new InvalidDataException($"Word numbering definition {id.Value} contains duplicate level {levelIndex}.");
                }
                while (definition.Levels.Count < levelIndex) {
                    definition.AddLevel();
                }
                RtfListLevel level = definition.AddLevel(format == NumberFormatValues.Bullet ? RtfListKind.Bullet : RtfListKind.Decimal);
                level.NumberFormat = ToRtfNumberFormat(format);
                level.StartAt = wordLevel.StartNumberingValue?.Val?.Value;
                level.Text = wordLevel.LevelText?.Val?.Value;
                level.LeftIndentTwips = ParseInt(wordLevel.PreviousParagraphProperties?.Indentation?.Left?.Value);
                level.FirstLineIndentTwips = ParseInt(wordLevel.PreviousParagraphProperties?.Indentation?.FirstLine?.Value)
                    ?? Negate(ParseInt(wordLevel.PreviousParagraphProperties?.Indentation?.Hanging?.Value));
                level.Alignment = ToRtfListAlignment(wordLevel.LevelJustification?.Val?.Value);
                level.FollowCharacter = ToRtfFollowCharacter(wordLevel.LevelSuffix?.Val?.Value);
            }
        }

        foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>()) {
            int? id = instance.NumberID?.Value;
            int? abstractId = instance.AbstractNumId?.Val?.Value;
            if (!id.HasValue || !abstractId.HasValue) continue;
            RtfListOverride item = destination.AddListOverride(id.Value, abstractId.Value);
            int overrideIndex = 0;
            foreach (LevelOverride wordOverride in instance.Elements<LevelOverride>()) {
                int effectiveLevelIndex = wordOverride.LevelIndex?.Value ?? overrideIndex;
                ValidateWordListLevel(effectiveLevelIndex);
                RtfListLevelOverride levelOverride = item.AddLevelOverride();
                levelOverride.LevelIndex = effectiveLevelIndex;
                levelOverride.StartAt = wordOverride.StartOverrideNumberingValue?.Val?.Value;
                levelOverride.OverrideStartAt = levelOverride.StartAt.HasValue;
                overrideIndex++;
            }
        }
    }

    private static void ApplyRtfStylesAndNumbering(RtfDocument source, WordDocument destination) {
        MainDocumentPart? main = destination._wordprocessingDocument.MainDocumentPart;
        if (main == null) return;
        StyleDefinitionsPart stylePart = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles ??= new Styles();
        var stylesById = new Dictionary<string, Style>(StringComparer.Ordinal);
        foreach (Style existingStyle in stylePart.Styles.Elements<Style>()) {
            string? existingId = existingStyle.StyleId?.Value;
            if (!string.IsNullOrEmpty(existingId)) stylesById[existingId!] = existingStyle;
        }
        foreach (RtfStyle style in source.Styles) {
            string wordStyleId = GetWordStyleId(style.Id, style.Kind);
            if (stylesById.TryGetValue(wordStyleId, out Style? existing)) existing.Remove();
            Style converted = CreateWordStyle(style, source);
            stylePart.Styles.Append(converted);
            stylesById[wordStyleId] = converted;
        }

        ApplyRtfNumbering(source, main);
    }

    private static Style CreateWordStyle(RtfStyle source, RtfDocument document) {
        var style = new Style { Type = ToWordStyleKind(source.Kind), StyleId = GetWordStyleId(source.Id, source.Kind) };
        style.Append(new StyleName { Val = source.Name });
        if (source.BasedOnStyleId.HasValue) style.Append(new BasedOn { Val = GetWordStyleId(source.BasedOnStyleId.Value, source.Kind) });
        if (source.NextStyleId.HasValue) style.Append(new NextParagraphStyle { Val = GetWordStyleId(source.NextStyleId.Value, RtfStyleKind.Paragraph) });
        if (source.LinkedStyleId.HasValue) style.Append(new LinkedStyle { Val = GetWordStyleId(source.LinkedStyleId.Value, source.Kind == RtfStyleKind.Paragraph ? RtfStyleKind.Character : RtfStyleKind.Paragraph) });
        if (source.AutoUpdate) style.Append(new AutoRedefine());
        if (source.Hidden) style.Append(new StyleHidden());
        if (source.Locked) style.Append(new Locked());
        if (source.SemiHidden) style.Append(new SemiHidden());
        if (source.UnhideWhenUsed) style.Append(new UnhideWhenUsed());
        if (source.QuickFormat) style.Append(new PrimaryStyle());
        if (source.Priority.HasValue) style.Append(new UIPriority { Val = source.Priority.Value });

        var run = new StyleRunProperties();
        if (source.Bold.HasValue) run.Bold = new Bold { Val = source.Bold.Value };
        if (source.Italic.HasValue) run.Italic = new Italic { Val = source.Italic.Value };
        if (source.UnderlineStyle.HasValue) run.Underline = new Underline { Val = ToWordUnderlineStyle(source.UnderlineStyle.Value) };
        if (source.FontSize.HasValue) run.FontSize = new FontSize { Val = (source.FontSize.Value * 2d).ToString("0.##", CultureInfo.InvariantCulture) };
        if (source.FontId.HasValue && TryGetFontName(document, source.FontId.Value, out string? fontName)) run.RunFonts = new RunFonts { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName };
        if (run.HasChildren) style.Append(run);

        var paragraph = new StyleParagraphProperties();
        if (source.ParagraphAlignment.HasValue) paragraph.Justification = new Justification { Val = ToWordTextAlignment(source.ParagraphAlignment.Value) };
        if (source.LeftIndentTwips.HasValue || source.RightIndentTwips.HasValue || source.FirstLineIndentTwips.HasValue) {
            paragraph.Indentation = new Indentation {
                Left = FormatInt(source.LeftIndentTwips),
                Right = FormatInt(source.RightIndentTwips),
                FirstLine = source.FirstLineIndentTwips >= 0 ? FormatInt(source.FirstLineIndentTwips) : null,
                Hanging = source.FirstLineIndentTwips < 0 ? FormatInt(-source.FirstLineIndentTwips) : null
            };
        }

        if (source.SpaceBeforeTwips.HasValue || source.SpaceAfterTwips.HasValue || source.LineSpacingTwips.HasValue) paragraph.SpacingBetweenLines = new SpacingBetweenLines { Before = FormatInt(source.SpaceBeforeTwips), After = FormatInt(source.SpaceAfterTwips), Line = FormatInt(source.LineSpacingTwips) };
        if (source.PageBreakBefore.HasValue) paragraph.PageBreakBefore = new PageBreakBefore { Val = source.PageBreakBefore.Value };
        if (source.KeepWithNext.HasValue) paragraph.KeepNext = new KeepNext { Val = source.KeepWithNext.Value };
        if (source.KeepLinesTogether.HasValue) paragraph.KeepLines = new KeepLines { Val = source.KeepLinesTogether.Value };
        if (source.OutlineLevel.HasValue) paragraph.OutlineLevel = new OutlineLevel { Val = source.OutlineLevel.Value };
        if (paragraph.HasChildren) style.Append(paragraph);
        return style;
    }

    private static void ApplyRtfNumbering(RtfDocument source, MainDocumentPart main) {
        RtfDocumentWriter.EffectiveListTables effectiveLists = RtfDocumentWriter.BuildEffectiveListTables(source);
        if (effectiveLists.Definitions.Count == 0 && effectiveLists.Overrides.Count == 0) return;
        NumberingDefinitionsPart part = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
        var numbering = new Numbering();
        foreach (RtfListDefinition definition in effectiveLists.Definitions) {
            var abstractNum = new AbstractNum { AbstractNumberId = definition.Id };
            IReadOnlyList<RtfListLevel> levels = definition.Levels.Count == 0
                ? new[] { new RtfListLevel(0, RtfListKind.Decimal) }
                : definition.Levels;
            foreach (RtfListLevel sourceLevel in levels) {
                var level = new Level { LevelIndex = sourceLevel.LevelIndex };
                level.Append(new StartNumberingValue { Val = sourceLevel.StartAt ?? 1 });
                level.Append(new NumberingFormat { Val = ToWordNumberFormat(sourceLevel) });
                level.Append(new LevelText { Val = string.IsNullOrEmpty(sourceLevel.Text) ? DefaultLevelText(sourceLevel) : sourceLevel.Text });
                if (sourceLevel.Alignment.HasValue) level.Append(new LevelJustification { Val = ToWordListAlignment(sourceLevel.Alignment.Value) });
                if (sourceLevel.FollowCharacter.HasValue) level.Append(new LevelSuffix { Val = ToWordFollowCharacter(sourceLevel.FollowCharacter.Value) });
                if (sourceLevel.LeftIndentTwips.HasValue || sourceLevel.FirstLineIndentTwips.HasValue) {
                    level.Append(new PreviousParagraphProperties(new Indentation {
                        Left = FormatInt(sourceLevel.LeftIndentTwips),
                        Hanging = sourceLevel.FirstLineIndentTwips < 0 ? FormatInt(-sourceLevel.FirstLineIndentTwips) : null,
                        FirstLine = sourceLevel.FirstLineIndentTwips >= 0 ? FormatInt(sourceLevel.FirstLineIndentTwips) : null
                    }));
                }

                abstractNum.Append(level);
            }

            numbering.Append(abstractNum);
        }

        foreach (RtfListOverride item in effectiveLists.Overrides) {
            var instance = new NumberingInstance { NumberID = item.Id };
            instance.Append(new AbstractNumId { Val = item.ListId });
            for (int levelIndex = 0; levelIndex < item.LevelOverrides.Count; levelIndex++) {
                RtfListLevelOverride sourceOverride = item.LevelOverrides[levelIndex];
                var levelOverride = new LevelOverride { LevelIndex = sourceOverride.LevelIndex ?? levelIndex };
                if (sourceOverride.OverrideStartAt == true && sourceOverride.StartAt.HasValue) {
                    levelOverride.Append(new StartOverrideNumberingValue { Val = sourceOverride.StartAt.Value });
                }
                instance.Append(levelOverride);
            }

            numbering.Append(instance);
        }

        part.Numbering = numbering;
    }

    private static void CopyParagraphStyleAndNumbering(WordParagraph source, RtfParagraph destination, RtfDocument document) {
        if (!string.IsNullOrWhiteSpace(source.StyleId)) destination.StyleId = FindRtfStyleId(document, source.StyleId!, RtfStyleKind.Paragraph);
        int? listId = source._listNumberId;
        if (!listId.HasValue) return;
        destination.ListId = listId;
        destination.ListLevel = source.ListItemLevel ?? 0;
        RtfListOverride? listOverride = document.ListOverrides.FirstOrDefault(item => item.Id == listId.Value);
        destination.ListDefinitionId = listOverride?.ListId;
        RtfListDefinition? definition = listOverride == null ? null : document.ListDefinitions.FirstOrDefault(item => item.Id == listOverride.ListId);
        RtfListLevel? level = definition?.Levels.FirstOrDefault(item => item.LevelIndex == destination.ListLevel);
        destination.ListKind = level?.Kind ?? RtfListKind.Decimal;
    }

    private static void ApplyParagraphStyleAndNumbering(WordParagraph destination, RtfParagraph source, RtfDocument document) {
        if (source.StyleId.HasValue && HasRtfStyle(document, source.StyleId.Value, RtfStyleKind.Paragraph)) destination.SetStyleId(GetWordStyleId(source.StyleId.Value, RtfStyleKind.Paragraph));
        if (!source.ListId.HasValue) return;
        ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
        properties.NumberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = source.ListLevel ?? 0 },
            new NumberingId { Val = source.ListId.Value });
    }

    private static int? FindRtfStyleId(RtfDocument document, string name, RtfStyleKind kind) =>
        document.Styles.FirstOrDefault(style => style.Kind == kind && string.Equals(style.Name, name, StringComparison.OrdinalIgnoreCase))?.Id;

    private static bool HasRtfStyle(RtfDocument document, int id, RtfStyleKind kind) =>
        document.Styles.Any(style => style.Id == id && style.Kind == kind);

    private static string GetWordStyleId(int id, RtfStyleKind kind) => "Rtf" + (kind == RtfStyleKind.Character ? "C" : kind == RtfStyleKind.Table ? "T" : "P") + id.ToString(CultureInfo.InvariantCulture);
    private static void ValidateWordListLevel(int levelIndex) {
        if (levelIndex < 0 || levelIndex > MaximumWordListLevel) {
            throw new InvalidDataException($"Word numbering level {levelIndex} is outside the supported range 0-{MaximumWordListLevel}.");
        }
    }
    private static int? ResolveStyleReference(string? id, IReadOnlyDictionary<string, int> ids) => id != null && ids.TryGetValue(id, out int value) ? value : null;
    private static RtfStyleKind ToRtfStyleKind(StyleValues? kind) => kind == StyleValues.Character ? RtfStyleKind.Character : kind == StyleValues.Table ? RtfStyleKind.Table : RtfStyleKind.Paragraph;
    private static StyleValues ToWordStyleKind(RtfStyleKind kind) => kind == RtfStyleKind.Character ? StyleValues.Character : kind == RtfStyleKind.Table ? StyleValues.Table : StyleValues.Paragraph;
    private static bool? ReadToggle(OnOffType? value) => value == null ? null : value.Val?.Value ?? true;
    private static int? ParseInt(string? value) => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : null;
    private static int? Negate(int? value) => value.HasValue ? -value.Value : null;
    private static string? FormatInt(int? value) => value?.ToString(CultureInfo.InvariantCulture);
    private static RtfTextAlignment? ToRtfTextAlignment(JustificationValues? value) => value == JustificationValues.Center ? RtfTextAlignment.Center : value == JustificationValues.Right ? RtfTextAlignment.Right : value == JustificationValues.Both ? RtfTextAlignment.Justify : value.HasValue ? RtfTextAlignment.Left : null;
    private static JustificationValues ToWordTextAlignment(RtfTextAlignment value) => value == RtfTextAlignment.Center ? JustificationValues.Center : value == RtfTextAlignment.Right ? JustificationValues.Right : value == RtfTextAlignment.Justify ? JustificationValues.Both : JustificationValues.Left;
    private static RtfListLevelAlignment? ToRtfListAlignment(LevelJustificationValues? value) => value == LevelJustificationValues.Center ? RtfListLevelAlignment.Center : value == LevelJustificationValues.Right ? RtfListLevelAlignment.Right : value.HasValue ? RtfListLevelAlignment.Left : null;
    private static LevelJustificationValues ToWordListAlignment(RtfListLevelAlignment value) => value == RtfListLevelAlignment.Center ? LevelJustificationValues.Center : value == RtfListLevelAlignment.Right ? LevelJustificationValues.Right : LevelJustificationValues.Left;
    private static RtfListLevelFollowCharacter? ToRtfFollowCharacter(LevelSuffixValues? value) => value == LevelSuffixValues.Space ? RtfListLevelFollowCharacter.Space : value == LevelSuffixValues.Nothing ? RtfListLevelFollowCharacter.Nothing : value.HasValue ? RtfListLevelFollowCharacter.Tab : null;
    private static LevelSuffixValues ToWordFollowCharacter(RtfListLevelFollowCharacter value) => value == RtfListLevelFollowCharacter.Space ? LevelSuffixValues.Space : value == RtfListLevelFollowCharacter.Nothing ? LevelSuffixValues.Nothing : LevelSuffixValues.Tab;
    private static int? ToRtfNumberFormat(NumberFormatValues? value) => value == NumberFormatValues.Bullet ? 23 : value == NumberFormatValues.UpperRoman ? 1 : value == NumberFormatValues.LowerRoman ? 2 : value == NumberFormatValues.UpperLetter ? 3 : value == NumberFormatValues.LowerLetter ? 4 : value.HasValue ? 0 : null;
    private static NumberFormatValues ToWordNumberFormat(RtfListLevel level) => level.Kind == RtfListKind.Bullet || level.NumberFormat == 23 ? NumberFormatValues.Bullet : level.NumberFormat == 1 ? NumberFormatValues.UpperRoman : level.NumberFormat == 2 ? NumberFormatValues.LowerRoman : level.NumberFormat == 3 ? NumberFormatValues.UpperLetter : level.NumberFormat == 4 ? NumberFormatValues.LowerLetter : NumberFormatValues.Decimal;
    private static string DefaultLevelText(RtfListLevel level) => level.Kind == RtfListKind.Bullet ? "•" : "%" + (level.LevelIndex + 1).ToString(CultureInfo.InvariantCulture) + ".";
}
