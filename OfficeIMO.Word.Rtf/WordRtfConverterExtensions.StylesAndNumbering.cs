using System.Globalization;
using OfficeIMO.Rtf.Writing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private const int MaximumWordListLevel = 8;

    private static void CopyWordStylesAndNumbering(WordDocument source, RtfDocument destination) {
        OfficeXmlNode styles = source.StylesNode();
        if (styles.Exists) {
            OfficeXmlNode[] definitions = styles.WElements("style").Where(style => style.WAttr("styleId") != null).ToArray();
            string? defaultParagraphId = definitions.LastOrDefault(style =>
                (style.WAttr("type") ?? "paragraph") == "paragraph" && WordXml.ParseOnOff(style.WAttr("default")) == true).WAttr("styleId");
            var ids = new Dictionary<string, int>(StringComparer.Ordinal);
            // RTF style zero is the implicit paragraph default, regardless of the Word catalogue's order.
            if (defaultParagraphId != null) ids[defaultParagraphId] = 0;
            int nextId = 1;
            foreach (OfficeXmlNode style in definitions) {
                string id = style.WAttr("styleId")!;
                if (!ids.ContainsKey(id)) ids[id] = nextId++;
            }

            foreach (OfficeXmlNode style in definitions) {
                string styleId = style.WAttr("styleId")!;
                RtfStyleKind kind = ToRtfStyleKind(style.WAttr("type"));
                RtfStyle target = destination.AddStyle(ids[styleId], styleId, kind);
                target.BasedOnStyleId = ResolveStyleReference(style.ChildVal("basedOn"), ids);
                target.NextStyleId = ResolveStyleReference(style.ChildVal("next"), ids);
                target.LinkedStyleId = ResolveStyleReference(style.ChildVal("link"), ids);
                target.AutoUpdate = style.WChild("autoRedefine").Exists;
                target.Hidden = style.WChild("hidden").Exists;
                target.Locked = style.WChild("locked").Exists;
                target.SemiHidden = style.WChild("semiHidden").Exists;
                target.UnhideWhenUsed = style.WChild("unhideWhenUsed").Exists;
                target.QuickFormat = style.WChild("qFormat").Exists;
                target.Priority = WordXml.ParseInt32(style.ChildVal("uiPriority"));
                CopyWordStyleFormatting(style, target, destination);
            }
        }

        CopyWordNumbering(source, destination);
    }

    private static void CopyWordStyleFormatting(OfficeXmlNode source, RtfStyle destination, RtfDocument document) {
        OfficeXmlNode run = source.WChild("rPr");
        if (run.Exists) {
            destination.Bold = ReadToggle(run.WChild("b"));
            destination.Italic = ReadToggle(run.WChild("i"));
            destination.TextHidden = ReadToggle(run.WChild("vanish"));
            if (ParseUnderlineToken(run.ChildVal("u")) is WordUnderlineStyle underline) destination.UnderlineStyle = ToRtfUnderlineStyle(underline);
            if (OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(run.ChildVal("sz"), NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints)) destination.FontSize = halfPoints / 2d;
            OfficeXmlNode runFonts = run.WChild("rFonts");
            string? font = runFonts.WAttr("ascii") ?? runFonts.WAttr("hAnsi") ?? runFonts.WAttr("eastAsia");
            if (!string.IsNullOrWhiteSpace(font)) destination.FontId = document.AddFont(font!);
            string? color = run.ChildVal("color");
            if (string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)) destination.ForegroundColorIndex = 0;
            else if (color != null && TryParseHexColor(color, out byte red, out byte green, out byte blue)) destination.ForegroundColorIndex = GetOrAddColor(document, red, green, blue);
            OfficeXmlNode characterShading = run.WChild("shd");
            string? fill = characterShading.WAttr("fill");
            string? pattern = characterShading.WVal();
            if (pattern == "clear" || pattern == "solid") {
                if (string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase)) destination.HighlightColorIndex = 0;
                else if (fill != null && TryParseHexColor(fill, out byte fillRed, out byte fillGreen, out byte fillBlue)) destination.HighlightColorIndex = GetOrAddColor(document, fillRed, fillGreen, fillBlue);
            }
        }

        OfficeXmlNode paragraph = source.WChild("pPr");
        if (paragraph.IsNone) return;
        destination.ParagraphAlignment = ToRtfTextAlignment(paragraph.ChildVal("jc"));
        OfficeXmlNode numberingProperties = paragraph.WChild("numPr");
        destination.ListId = WordXml.ParseInt32(numberingProperties.ChildVal("numId"));
        destination.ListLevel = WordXml.ParseInt32(numberingProperties.ChildVal("ilvl"));
        OfficeXmlNode indentation = paragraph.WChild("ind");
        destination.LeftIndentTwips = ParseInt(indentation.WAttr("left"));
        destination.RightIndentTwips = ParseInt(indentation.WAttr("right"));
        destination.FirstLineIndentTwips = ParseInt(indentation.WAttr("firstLine")) ?? Negate(ParseInt(indentation.WAttr("hanging")));
        OfficeXmlNode spacing = paragraph.WChild("spacing");
        destination.SpaceBeforeTwips = ParseInt(spacing.WAttr("before"));
        destination.SpaceAfterTwips = ParseInt(spacing.WAttr("after"));
        int? lineSpacing = ParseInt(spacing.WAttr("line"));
        string? lineSpacingRule = spacing.WAttr("lineRule");
        destination.LineSpacingTwips = ToRtfLineSpacingValue(lineSpacing, lineSpacingRule);
        destination.LineSpacingMultiple = ToRtfLineSpacingMultiple(lineSpacing, lineSpacingRule);
        OfficeXmlNode tabs = paragraph.WChild("tabs");
        if (tabs.Exists) {
            foreach (OfficeXmlNode tabStop in tabs.WElements("tab")) {
                string? alignment = tabStop.WVal();
                if (alignment == "clear") continue;
                if (WordXml.ParseInt32(tabStop.WAttr("pos")) is int position && position >= 0) {
                    destination.AddTabStop(position, ToRtfTabAlignment(alignment ?? "left"),
                        ToRtfTabLeader(tabStop.WAttr("leader") ?? "none"));
                }
            }
        }
        destination.PageBreakBefore = ReadToggle(paragraph.WChild("pageBreakBefore"));
        destination.KeepWithNext = ReadToggle(paragraph.WChild("keepNext"));
        destination.KeepLinesTogether = ReadToggle(paragraph.WChild("keepLines"));
        destination.OutlineLevel = WordXml.ParseInt32(paragraph.ChildVal("outlineLvl"));
    }

    private static void CopyWordNumbering(WordDocument source, RtfDocument destination) {
        OfficeXmlNode numbering = source.NumberingNode();
        if (numbering.IsNone) return;
        foreach (OfficeXmlNode abstractNum in numbering.WElements("abstractNum")) {
            int? id = WordXml.ParseInt32(abstractNum.WAttr("abstractNumId"));
            if (!id.HasValue) continue;
            RtfListDefinition definition = destination.AddListDefinition(id.Value);
            var seenLevels = new HashSet<int>();
            foreach (OfficeXmlNode wordLevel in abstractNum.WElements("lvl").OrderBy(level => WordXml.ParseInt32(level.WAttr("ilvl")) ?? 0)) {
                int levelIndex = WordXml.ParseInt32(wordLevel.WAttr("ilvl")) ?? definition.Levels.Count;
                ValidateWordListLevel(levelIndex);
                if (!seenLevels.Add(levelIndex)) {
                    throw new InvalidDataException($"Word numbering definition {id.Value} contains duplicate level {levelIndex}.");
                }
                while (definition.Levels.Count < levelIndex) {
                    definition.AddLevel();
                }
                definition.AddParsedLevel(ReadWordListLevel(wordLevel, levelIndex));
            }
        }

        foreach (OfficeXmlNode instance in numbering.WElements("num")) {
            int? id = WordXml.ParseInt32(instance.WAttr("numId"));
            int? abstractId = WordXml.ParseInt32(instance.ChildVal("abstractNumId"));
            if (!id.HasValue || !abstractId.HasValue) continue;
            RtfListOverride item = destination.AddListOverride(id.Value, abstractId.Value);
            int overrideIndex = 0;
            foreach (OfficeXmlNode wordOverride in instance.WElements("lvlOverride")) {
                int effectiveLevelIndex = WordXml.ParseInt32(wordOverride.WAttr("ilvl")) ?? overrideIndex;
                ValidateWordListLevel(effectiveLevelIndex);
                RtfListLevelOverride levelOverride = item.AddLevelOverride();
                levelOverride.LevelIndex = effectiveLevelIndex;
                levelOverride.StartAt = WordXml.ParseInt32(wordOverride.ChildVal("startOverride"));
                levelOverride.OverrideStartAt = levelOverride.StartAt.HasValue;
                OfficeXmlNode overrideLevel = wordOverride.WChild("lvl");
                if (overrideLevel.Exists) {
                    levelOverride.OverrideFormat = true;
                    levelOverride.Formatting = ReadWordListLevel(overrideLevel, effectiveLevelIndex);
                }
                overrideIndex++;
            }
        }
    }

    private static RtfListLevel ReadWordListLevel(OfficeXmlNode source, int index) {
        string? format = source.ChildVal("numFmt");
        OfficeXmlNode indentation = source.WChild("pPr").WChild("ind");
        return new RtfListLevel(index, format == "bullet" ? RtfListKind.Bullet : RtfListKind.Decimal) {
            NumberFormat = ToRtfNumberFormat(format),
            StartAt = WordXml.ParseInt32(source.ChildVal("start")),
            Text = source.ChildVal("lvlText"),
            LeftIndentTwips = ParseInt(indentation.WAttr("left")),
            FirstLineIndentTwips = ParseInt(indentation.WAttr("firstLine"))
                ?? Negate(ParseInt(indentation.WAttr("hanging"))),
            Alignment = ToRtfListAlignment(WordListLevel.ReadLevelJustification(source.WChild("lvlJc").WVal())),
            FollowCharacter = ToRtfFollowCharacter(source.ChildVal("suff")),
            NoRestart = WordXml.ParseInt32(source.ChildVal("lvlRestart")) == 0 ? true : (bool?)null,
            LegalNumbering = ReadToggle(source.WChild("isLgl"))
        };
    }

    /// <summary>A new <c>w:lvl</c> made for <paramref name="container"/> (a <c>w:abstractNum</c> or <c>w:lvlOverride</c>), not yet inserted.</summary>
    private static OfficeXmlNode CreateWordListLevel(OfficeXmlNode container, RtfListLevel source) {
        OfficeXmlNode level = container.CreateElement(WordXml.W, "lvl");
        level.SetAttribute(WordXml.W, "ilvl", FormatInt(source.LevelIndex));
        AppendValueChild(level, "start", FormatInt(source.StartAt ?? 1)!);
        AppendValueChild(level, "numFmt", ToWordNumberFormat(source));
        if (source.NoRestart == true) AppendValueChild(level, "lvlRestart", "0");
        if (source.LegalNumbering.HasValue) AppendValueChild(level, "isLgl", FormatOnOff(source.LegalNumbering.Value));
        if (source.FollowCharacter.HasValue) AppendValueChild(level, "suff", ToWordFollowCharacter(source.FollowCharacter.Value));
        AppendValueChild(level, "lvlText", source.Text ?? DefaultLevelText(source));
        if (source.Alignment.HasValue) AppendValueChild(level, "lvlJc", ToWordListAlignment(source.Alignment.Value));
        if (source.LeftIndentTwips.HasValue || source.FirstLineIndentTwips.HasValue) {
            OfficeXmlNode properties = level.CreateElement(WordXml.W, "pPr");
            OfficeXmlNode indentation = properties.CreateElement(WordXml.W, "ind");
            indentation.SetAttribute(WordXml.W, "left", FormatInt(source.LeftIndentTwips));
            indentation.SetAttribute(WordXml.W, "hanging", source.FirstLineIndentTwips < 0 ? FormatInt(-source.FirstLineIndentTwips) : null);
            indentation.SetAttribute(WordXml.W, "firstLine", source.FirstLineIndentTwips >= 0 ? FormatInt(source.FirstLineIndentTwips) : null);
            properties.Append(indentation);
            level.Append(properties);
        }
        return level;
    }

    private static void ApplyRtfStylesAndNumbering(RtfDocument source, WordDocument destination) {
        OfficeXmlNode styles = destination.GetOrAddStylesNode();
        if (styles.IsNone) return;
        var stylesById = new Dictionary<string, OfficeXmlNode>(StringComparer.Ordinal);
        foreach (OfficeXmlNode existingStyle in styles.WElements("style").ToList()) {
            string? existingId = existingStyle.WAttr("styleId");
            if (!string.IsNullOrEmpty(existingId)) stylesById[existingId!] = existingStyle;
        }
        foreach (RtfStyle style in source.Styles) {
            string wordStyleId = GetWordStyleId(style.Id, style.Kind);
            if (stylesById.TryGetValue(wordStyleId, out OfficeXmlNode existing)) existing.Remove();
            OfficeXmlNode converted = styles.Append(CreateWordStyle(styles, style, source));
            stylesById[wordStyleId] = converted;
        }

        ApplyRtfNumbering(source, destination);
    }

    private static OfficeXmlNode CreateWordStyle(OfficeXmlNode styles, RtfStyle source, RtfDocument document) {
        OfficeXmlNode style = styles.CreateElement(WordXml.W, "style");
        style.SetAttribute(WordXml.W, "type", ToWordStyleKind(source.Kind));
        style.SetAttribute(WordXml.W, "styleId", GetWordStyleId(source.Id, source.Kind));
        style.AddOrdered(WordXml.W, "name", WordXml.StyleOrder).SetAttribute(WordXml.W, "val", source.Name);
        if (source.BasedOnStyleId.HasValue) style.AddOrdered(WordXml.W, "basedOn", WordXml.StyleOrder).SetAttribute(WordXml.W, "val", GetWordStyleId(source.BasedOnStyleId.Value, source.Kind));
        if (source.NextStyleId.HasValue) style.AddOrdered(WordXml.W, "next", WordXml.StyleOrder).SetAttribute(WordXml.W, "val", GetWordStyleId(source.NextStyleId.Value, RtfStyleKind.Paragraph));
        if (source.LinkedStyleId.HasValue) style.AddOrdered(WordXml.W, "link", WordXml.StyleOrder).SetAttribute(WordXml.W, "val", GetWordStyleId(source.LinkedStyleId.Value, source.Kind == RtfStyleKind.Paragraph ? RtfStyleKind.Character : RtfStyleKind.Paragraph));
        if (source.AutoUpdate) style.AddOrdered(WordXml.W, "autoRedefine", WordXml.StyleOrder);
        if (source.Hidden) style.AddOrdered(WordXml.W, "hidden", WordXml.StyleOrder);
        if (source.Locked) style.AddOrdered(WordXml.W, "locked", WordXml.StyleOrder);
        if (source.SemiHidden) style.AddOrdered(WordXml.W, "semiHidden", WordXml.StyleOrder);
        if (source.UnhideWhenUsed) style.AddOrdered(WordXml.W, "unhideWhenUsed", WordXml.StyleOrder);
        if (source.QuickFormat) style.AddOrdered(WordXml.W, "qFormat", WordXml.StyleOrder);
        if (source.Priority.HasValue) style.AddOrdered(WordXml.W, "uiPriority", WordXml.StyleOrder).SetAttribute(WordXml.W, "val", FormatInt(source.Priority.Value));

        OfficeXmlNode run = style.CreateElement(WordXml.W, "rPr");
        if (source.Bold.HasValue) run.AddOrdered(WordXml.W, "b", WRPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(source.Bold.Value));
        if (source.Italic.HasValue) run.AddOrdered(WordXml.W, "i", WRPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(source.Italic.Value));
        if (source.TextHidden.HasValue) run.AddOrdered(WordXml.W, "vanish", WRPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(source.TextHidden.Value));
        if (source.UnderlineStyle.HasValue) run.AddOrdered(WordXml.W, "u", WRPrView.Order).SetAttribute(WordXml.W, "val", ToUnderlineToken(ToWordUnderlineStyle(source.UnderlineStyle.Value)));
        if (source.FontSize.HasValue) run.AddOrdered(WordXml.W, "sz", WRPrView.Order).SetAttribute(WordXml.W, "val", (source.FontSize.Value * 2d).ToString("0.##", CultureInfo.InvariantCulture));
        if (source.FontId.HasValue && TryGetFontName(document, source.FontId.Value, out string? fontName)) {
            OfficeXmlNode fonts = run.AddOrdered(WordXml.W, "rFonts", WRPrView.Order);
            fonts.SetAttribute(WordXml.W, "ascii", fontName);
            fonts.SetAttribute(WordXml.W, "hAnsi", fontName);
            fonts.SetAttribute(WordXml.W, "eastAsia", fontName);
            fonts.SetAttribute(WordXml.W, "cs", fontName);
        }
        if (source.ForegroundColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ForegroundColorIndex.Value);
            if (color != null || source.ForegroundColorIndex == 0) run.AddOrdered(WordXml.W, "color", WRPrView.Order).SetAttribute(WordXml.W, "val", color ?? "auto");
        }
        if (source.HighlightColorIndex.HasValue) {
            string? fill = GetColorHex(document, source.HighlightColorIndex.Value);
            if (fill != null || source.HighlightColorIndex == 0) {
                OfficeXmlNode shading = run.AddOrdered(WordXml.W, "shd", WRPrView.Order);
                shading.SetAttribute(WordXml.W, "val", "clear");
                shading.SetAttribute(WordXml.W, "fill", fill ?? "auto");
            }
        }
        if (run.HasChildren()) style.InsertOrdered(run, WordXml.StyleOrder);

        OfficeXmlNode paragraph = style.CreateElement(WordXml.W, "pPr");
        if (source.ListId.HasValue) {
            OfficeXmlNode numberingProperties = paragraph.AddOrdered(WordXml.W, "numPr", WPPrView.Order);
            AppendValueChild(numberingProperties, "ilvl", FormatInt(source.ListLevel ?? 0)!);
            AppendValueChild(numberingProperties, "numId", FormatInt(source.ListId.Value)!);
        }
        if (source.ParagraphAlignment.HasValue) paragraph.AddOrdered(WordXml.W, "jc", WPPrView.Order).SetAttribute(WordXml.W, "val", ToWordTextAlignment(source.ParagraphAlignment.Value));
        if (source.LeftIndentTwips.HasValue || source.RightIndentTwips.HasValue || source.FirstLineIndentTwips.HasValue) {
            OfficeXmlNode indentation = paragraph.AddOrdered(WordXml.W, "ind", WPPrView.Order);
            indentation.SetAttribute(WordXml.W, "left", FormatInt(source.LeftIndentTwips));
            indentation.SetAttribute(WordXml.W, "right", FormatInt(source.RightIndentTwips));
            indentation.SetAttribute(WordXml.W, "hanging", source.FirstLineIndentTwips < 0 ? FormatInt(-source.FirstLineIndentTwips) : null);
            indentation.SetAttribute(WordXml.W, "firstLine", source.FirstLineIndentTwips >= 0 ? FormatInt(source.FirstLineIndentTwips) : null);
        }

        if (source.SpaceBeforeTwips.HasValue || source.SpaceAfterTwips.HasValue || source.LineSpacingTwips.HasValue) {
            OfficeXmlNode spacing = paragraph.AddOrdered(WordXml.W, "spacing", WPPrView.Order);
            spacing.SetAttribute(WordXml.W, "before", FormatInt(source.SpaceBeforeTwips));
            spacing.SetAttribute(WordXml.W, "after", FormatInt(source.SpaceAfterTwips));
            spacing.SetAttribute(WordXml.W, "line", FormatInt(ToWordLineSpacingValue(source.LineSpacingTwips, source.LineSpacingMultiple)));
            spacing.SetAttribute(WordXml.W, "lineRule", ToWordLineSpacingRule(source.LineSpacingTwips, source.LineSpacingMultiple));
        }
        if (source.TabStops.Count > 0) {
            OfficeXmlNode tabs = paragraph.AddOrdered(WordXml.W, "tabs", WPPrView.Order);
            foreach (RtfTabStop tabStop in source.TabStops) {
                OfficeXmlNode tab = tabs.CreateElement(WordXml.W, "tab");
                tab.SetAttribute(WordXml.W, "val", ToWordTabAlignment(tabStop.Alignment).ToOpenXmlToken());
                tab.SetAttribute(WordXml.W, "leader", ToWordTabLeader(tabStop.Leader).ToOpenXmlToken());
                tab.SetAttribute(WordXml.W, "pos", FormatInt(tabStop.PositionTwips));
                tabs.Append(tab);
            }
        }
        if (source.PageBreakBefore.HasValue) paragraph.AddOrdered(WordXml.W, "pageBreakBefore", WPPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(source.PageBreakBefore.Value));
        if (source.KeepWithNext.HasValue) paragraph.AddOrdered(WordXml.W, "keepNext", WPPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(source.KeepWithNext.Value));
        if (source.KeepLinesTogether.HasValue) paragraph.AddOrdered(WordXml.W, "keepLines", WPPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(source.KeepLinesTogether.Value));
        if (source.OutlineLevel.HasValue) paragraph.AddOrdered(WordXml.W, "outlineLvl", WPPrView.Order).SetAttribute(WordXml.W, "val", FormatInt(source.OutlineLevel.Value));
        if (paragraph.HasChildren()) style.InsertOrdered(paragraph, WordXml.StyleOrder);
        return style;
    }

    private static void ApplyRtfNumbering(RtfDocument source, WordDocument destination) {
        RtfDocumentWriter.EffectiveListTables effectiveLists = RtfDocumentWriter.BuildEffectiveListTables(source);
        if (effectiveLists.Definitions.Count == 0 && effectiveLists.Overrides.Count == 0) return;
        OfficeXmlNode numbering = WordConverterNodes.CreateDetachedNumbering();
        foreach (RtfListDefinition definition in effectiveLists.Definitions) {
            OfficeXmlNode abstractNum = numbering.CreateElement(WordXml.W, "abstractNum");
            abstractNum.SetAttribute(WordXml.W, "abstractNumId", FormatInt(definition.Id));
            IReadOnlyList<RtfListLevel> levels = definition.Levels.Count == 0
                ? new[] { new RtfListLevel(0, RtfListKind.Decimal) }
                : definition.Levels;
            foreach (RtfListLevel sourceLevel in levels) {
                abstractNum.Append(CreateWordListLevel(abstractNum, sourceLevel));
            }

            numbering.Append(abstractNum);
        }

        foreach (RtfListOverride item in effectiveLists.Overrides) {
            OfficeXmlNode instance = numbering.CreateElement(WordXml.W, "num");
            instance.SetAttribute(WordXml.W, "numId", FormatInt(item.Id));
            AppendValueChild(instance, "abstractNumId", FormatInt(item.ListId)!);
            for (int levelIndex = 0; levelIndex < item.LevelOverrides.Count; levelIndex++) {
                RtfListLevelOverride sourceOverride = item.LevelOverrides[levelIndex];
                OfficeXmlNode levelOverride = instance.CreateElement(WordXml.W, "lvlOverride");
                levelOverride.SetAttribute(WordXml.W, "ilvl", FormatInt(sourceOverride.LevelIndex ?? levelIndex));
                if (sourceOverride.OverrideStartAt == true && sourceOverride.StartAt.HasValue) {
                    AppendValueChild(levelOverride, "startOverride", FormatInt(sourceOverride.StartAt.Value)!);
                }
                if (sourceOverride.OverrideFormat == true && sourceOverride.Formatting != null) {
                    RtfListLevel formatting = source.ResolveListFormatting(new RtfParagraph { ListId = item.Id, ListLevel = sourceOverride.LevelIndex ?? levelIndex })!.Level;
                    levelOverride.Append(CreateWordListLevel(levelOverride, formatting));
                }
                instance.Append(levelOverride);
            }

            numbering.Append(instance);
        }

        destination.SetNumberingNode(numbering);
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

    private static void ApplyParagraphStyleAndNumbering(WordParagraph destination, RtfParagraph source, RtfDocument document, int? projectedNumberId) {
        if (source.StyleId.HasValue && HasRtfStyle(document, source.StyleId.Value, RtfStyleKind.Paragraph)) destination.SetStyleId(GetWordStyleId(source.StyleId.Value, RtfStyleKind.Paragraph));
        int? listId = projectedNumberId ?? source.ListId;
        if (!listId.HasValue) return;
        OfficeXmlNode numberingProperties = WordXmlEdits.ReplaceChild(GetOrAddParagraphProperties(destination), "numPr", WPPrView.Order);
        AppendValueChild(numberingProperties, "ilvl", FormatInt(source.ListLevel ?? 0)!);
        AppendValueChild(numberingProperties, "numId", FormatInt(listId.Value)!);
    }

    private static int? FindRtfStyleId(RtfDocument document, string name, RtfStyleKind kind) =>
        document.Styles.FirstOrDefault(style => style.Kind == kind && string.Equals(style.Name, name, StringComparison.Ordinal))?.Id;

    private static bool HasRtfStyle(RtfDocument document, int id, RtfStyleKind kind) =>
        document.Styles.Any(style => style.Id == id && style.Kind == kind);

    private static string GetWordStyleId(int id, RtfStyleKind kind) => "Rtf" + (kind == RtfStyleKind.Character ? "C" : kind == RtfStyleKind.Table ? "T" : "P") + id.ToString(CultureInfo.InvariantCulture);
    private static void ValidateWordListLevel(int levelIndex) {
        if (levelIndex < 0 || levelIndex > MaximumWordListLevel) {
            throw new InvalidDataException($"Word numbering level {levelIndex} is outside the supported range 0-{MaximumWordListLevel}.");
        }
    }
    private static int? ResolveStyleReference(string? id, IReadOnlyDictionary<string, int> ids) => id != null && ids.TryGetValue(id, out int value) ? value : null;
    // Style kinds are w:type tokens.
    private static RtfStyleKind ToRtfStyleKind(string? kind) => kind == "character" ? RtfStyleKind.Character : kind == "table" ? RtfStyleKind.Table : RtfStyleKind.Paragraph;
    private static string ToWordStyleKind(RtfStyleKind kind) => kind == RtfStyleKind.Character ? "character" : kind == RtfStyleKind.Table ? "table" : "paragraph";
    /// <summary>An on/off property as the SDK's <c>OnOffType</c> reads it: null when missing, on when it has no value.</summary>
    private static bool? ReadToggle(OfficeXmlNode value) => value.IsNone ? null : value.IsOn();
    private static int? ParseInt(string? value) => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : null;
    private static int? Negate(int? value) => value.HasValue ? -value.Value : null;
    private static string? FormatInt(int? value) => value?.ToString(CultureInfo.InvariantCulture);
    private static string FormatOnOff(bool value) => value ? "true" : "false";
    // Alignments are w:jc tokens.
    private static RtfTextAlignment? ToRtfTextAlignment(string? value) => value == "center" ? RtfTextAlignment.Center : value == "right" ? RtfTextAlignment.Right : value == "both" ? RtfTextAlignment.Justify : value != null ? RtfTextAlignment.Left : null;
    private static string ToWordTextAlignment(RtfTextAlignment value) => value == RtfTextAlignment.Center ? "center" : value == RtfTextAlignment.Right ? "right" : value == RtfTextAlignment.Justify ? "both" : "left";
    private static RtfListLevelAlignment? ToRtfListAlignment(WordListLevelAlignment? value) => value switch {
        WordListLevelAlignment.Center => RtfListLevelAlignment.Center,
        WordListLevelAlignment.Right => RtfListLevelAlignment.Right,
        WordListLevelAlignment.Left => RtfListLevelAlignment.Left,
        _ => null
    };
    private static string ToWordListAlignment(RtfListLevelAlignment value) => value == RtfListLevelAlignment.Center ? "center" : value == RtfListLevelAlignment.Right ? "right" : "left";
    // Follow characters are w:suff tokens.
    private static RtfListLevelFollowCharacter? ToRtfFollowCharacter(string? value) => value == "space" ? RtfListLevelFollowCharacter.Space : value == "nothing" ? RtfListLevelFollowCharacter.Nothing : value != null ? RtfListLevelFollowCharacter.Tab : null;
    private static string ToWordFollowCharacter(RtfListLevelFollowCharacter value) => value == RtfListLevelFollowCharacter.Space ? "space" : value == RtfListLevelFollowCharacter.Nothing ? "nothing" : "tab";
    // Number formats are w:numFmt tokens.
    private static int? ToRtfNumberFormat(string? value) => value == "bullet" ? 23 : value == "upperRoman" ? 1 : value == "lowerRoman" ? 2 : value == "upperLetter" ? 3 : value == "lowerLetter" ? 4 : value != null ? 0 : null;
    private static string ToWordNumberFormat(RtfListLevel level) => level.Kind == RtfListKind.Bullet || level.NumberFormat == 23 ? "bullet" : level.NumberFormat == 1 ? "upperRoman" : level.NumberFormat == 2 ? "lowerRoman" : level.NumberFormat == 3 ? "upperLetter" : level.NumberFormat == 4 ? "lowerLetter" : "decimal";
    private static string DefaultLevelText(RtfListLevel level) => level.Kind == RtfListKind.Bullet ? "•" : "%" + (level.LevelIndex + 1).ToString(CultureInfo.InvariantCulture) + ".";
}
