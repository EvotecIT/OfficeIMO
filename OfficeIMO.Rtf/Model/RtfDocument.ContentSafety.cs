using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    /// <summary>Inspects native RTF visibility, revision, geometry, contrast, and non-primary story evidence.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] rtf,
        OfficeContentSafetyOptions? options = null,
        RtfReadOptions? readOptions = null) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(rtf);
#else
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));
#endif
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateBytes(rtf, effective);
        RtfDocument document = Load(rtf, readOptions).Document;
        return InspectRtfContentSafety(document, effective, targets: null);
    }

    /// <summary>Inspects an RTF file without treating concealment as evidence of AI authorship.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null,
        RtfReadOptions? readOptions = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return InspectContentSafety(OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective), effective, readOptions);
    }

    /// <summary>Removes exact selected RTF run payloads and verifies the normalized rewritten artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] rtf,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        RtfReadOptions? readOptions = null,
        RtfWriteOptions? writeOptions = null) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(rtf);
        ArgumentNullException.ThrowIfNull(selection);
#else
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
#endif
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentSafetyReport before = InspectContentSafety(rtf, options.Inspection, readOptions);
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) return new OfficeContentCleanupResult((byte[])rtf.Clone(), before, before, Array.Empty<OfficeContentCleanupChange>());

        RtfDocument document = Load(rtf, readOptions).Document;
        var targets = new Dictionary<string, RtfContentSafetyTarget>(StringComparer.Ordinal);
        OfficeContentSafetyReport current = InspectRtfContentSafety(document, options.Inspection, targets);
        IReadOnlyList<OfficeContentSafetyFinding> currentSelection = OfficeContentSafetyBuilder.ResolveSelection(current, selection);
        foreach (OfficeContentSafetyFinding finding in currentSelection.OrderByDescending(item => item.SourceTextOffset ?? -1)) targets[finding.Id].Remove();

        byte[] output = document.ToBytes(writeOptions);
        OfficeContentSafetyReport after = InspectContentSafety(output, options.Inspection, readOptions);
        OfficeContentCleanupChange[] changes = selected
            .Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability))
            .ToArray();
        return new OfficeContentCleanupResult(output, before, after, changes);
    }

    /// <summary>Atomically writes an explicitly cleaned RTF artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        RtfReadOptions? readOptions = null,
        RtfWriteOptions? writeOptions = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentCleanupResult result = RemoveSelectedContent(OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection), selection, options, readOptions, writeOptions);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static OfficeContentSafetyReport InspectRtfContentSafety(
        RtfDocument document,
        OfficeContentSafetyOptions? options,
        IDictionary<string, RtfContentSafetyTarget>? targets) {
        var builder = new OfficeContentSafetyBuilder("RTF", options);
        if (document.Sections.Count > 0) {
            for (int index = 0; index < document.Sections.Count; index++) {
                InspectRtfBlocks(document, document.Sections[index].Blocks, "Section[" + (index + 1).ToString(CultureInfo.InvariantCulture) + "]", false, builder, targets);
            }
        } else {
            InspectRtfBlocks(document, document.Blocks, "Document", false, builder, targets);
        }
        for (int index = 0; index < document.HeaderFooters.Count; index++) {
            InspectRtfParagraphs(document, document.HeaderFooters[index].Paragraphs, "HeaderFooter[" + (index + 1).ToString(CultureInfo.InvariantCulture) + "]", false, null, builder, targets);
        }
        for (int index = 0; index < document.Notes.Count; index++) {
            InspectRtfParagraphs(document, document.Notes[index].Paragraphs, "Note[" + (index + 1).ToString(CultureInfo.InvariantCulture) + "]", true, null, builder, targets);
        }
        if (document.HtmlEncapsulation != null && !string.IsNullOrWhiteSpace(document.HtmlEncapsulation.Html) && builder.Options.IncludeNonPrimaryContent) {
            OfficeContentSafetyFinding finding = builder.Add(
                OfficeContentConcealmentKind.NonPrimaryContent,
                OfficeContentSafetyRisk.ContextDependent,
                "HtmlEncapsulation",
                "RTF HTML encapsulation is a separate machine-readable representation outside the ordinary RTF text story.",
                document.HtmlEncapsulation.Html,
                OfficeContentCleanupCapability.RemoveElement,
                inspectTextIntegrityEvidence: false);
            if (targets != null) targets[finding.Id] = RtfContentSafetyTarget.ForHtmlEncapsulation(document);
            IReadOnlyList<OfficeContentSafetyFinding> unicode = builder.InspectChargedTextIntegrity(
                "HtmlEncapsulation/Text",
                document.HtmlEncapsulation.Html,
                OfficeContentCleanupCapability.RemoveText);
            if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = RtfContentSafetyTarget.ForHtmlEncapsulationRange(document, item);
        }
        return builder.Build();
    }

    private static void InspectRtfBlocks(
        RtfDocument document,
        IReadOnlyList<IRtfBlock> blocks,
        string root,
        bool nonPrimary,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, RtfContentSafetyTarget>? targets) {
        for (int blockIndex = 0; blockIndex < blocks.Count; blockIndex++) {
            string location = root + "/Block[" + (blockIndex + 1).ToString(CultureInfo.InvariantCulture) + "]";
            switch (blocks[blockIndex]) {
                case RtfParagraph paragraph:
                    InspectRtfParagraph(document, paragraph, location, nonPrimary, null, builder, targets);
                    break;
                case RtfTable table:
                    for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                        RtfTableRow row = table.Rows[rowIndex];
                        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                            RtfTableCell cell = row.Cells[cellIndex];
                            string cellLocation = location + "/Row[" + (rowIndex + 1).ToString(CultureInfo.InvariantCulture) + "]/Cell[" + (cellIndex + 1).ToString(CultureInfo.InvariantCulture) + "]";
                            InspectRtfBlocks(document, cell.Blocks, cellLocation, nonPrimary, cell.BackgroundColorIndex, builder, targets);
                        }
                    }
                    break;
                case RtfShape shape:
                    InspectRtfParagraphs(document, shape.TextBoxParagraphs, location + "/ShapeText", nonPrimary, null, builder, targets);
                    break;
                case RtfObject embedded:
                    InspectRtfParagraph(document, embedded.Result, location + "/ObjectResult", true, null, builder, targets);
                    break;
            }
        }
    }

    private static void InspectRtfBlocks(
        RtfDocument document,
        IReadOnlyList<IRtfBlock> blocks,
        string root,
        bool nonPrimary,
        int? inheritedBackgroundColorIndex,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, RtfContentSafetyTarget>? targets) {
        for (int index = 0; index < blocks.Count; index++) {
            if (blocks[index] is RtfParagraph paragraph) {
                InspectRtfParagraph(document, paragraph, root + "/Block[" + (index + 1).ToString(CultureInfo.InvariantCulture) + "]", nonPrimary, inheritedBackgroundColorIndex, builder, targets);
            } else if (blocks[index] is RtfTable nested) {
                InspectRtfBlocks(document, new IRtfBlock[] { nested }, root + "/Nested", nonPrimary, builder, targets);
            } else if (blocks[index] is RtfShape shape) {
                InspectRtfParagraphs(document, shape.TextBoxParagraphs, root + "/ShapeText[" + (index + 1).ToString(CultureInfo.InvariantCulture) + "]", nonPrimary, inheritedBackgroundColorIndex, builder, targets);
            } else if (blocks[index] is RtfObject embedded) {
                InspectRtfParagraph(document, embedded.Result, root + "/ObjectResult[" + (index + 1).ToString(CultureInfo.InvariantCulture) + "]", true, inheritedBackgroundColorIndex, builder, targets);
            }
        }
    }

    private static void InspectRtfParagraphs(
        RtfDocument document,
        IReadOnlyList<RtfParagraph> paragraphs,
        string root,
        bool nonPrimary,
        int? inheritedBackgroundColorIndex,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, RtfContentSafetyTarget>? targets) {
        for (int index = 0; index < paragraphs.Count; index++) {
            InspectRtfParagraph(document, paragraphs[index], root + "/Paragraph[" + (index + 1).ToString(CultureInfo.InvariantCulture) + "]", nonPrimary, inheritedBackgroundColorIndex, builder, targets);
        }
    }

    private static void InspectRtfParagraph(
        RtfDocument document,
        RtfParagraph paragraph,
        string location,
        bool nonPrimary,
        int? inheritedBackgroundColorIndex,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, RtfContentSafetyTarget>? targets) {
        int? paragraphBackground = paragraph.BackgroundColorIndex ?? ResolveRtfParagraphStyleBackground(document, paragraph.StyleId) ?? inheritedBackgroundColorIndex;
        int runIndex = 0;
        int inlineIndex = 0;
        foreach (IRtfInline inline in paragraph.Inlines) {
            inlineIndex++;
            if (inline is RtfRun run) {
                InspectRtfRun(document, paragraph, run, location + "/Run[" + (++runIndex).ToString(CultureInfo.InvariantCulture) + "]", nonPrimary, paragraphBackground, builder, targets);
            } else if (inline is RtfField field) {
                InspectRtfParagraph(document, field.Result, location + "/Field[" + inlineIndex.ToString(CultureInfo.InvariantCulture) + "]/Result", nonPrimary, paragraphBackground, builder, targets);
            } else if (inline is RtfShape shape) {
                InspectRtfParagraphs(document, shape.TextBoxParagraphs, location + "/Shape[" + inlineIndex.ToString(CultureInfo.InvariantCulture) + "]", nonPrimary, paragraphBackground, builder, targets);
            } else if (inline is RtfObject embedded) {
                InspectRtfParagraph(document, embedded.Result, location + "/Object[" + inlineIndex.ToString(CultureInfo.InvariantCulture) + "]/Result", true, paragraphBackground, builder, targets);
            }
        }
    }

    private static void InspectRtfRun(
        RtfDocument document,
        RtfParagraph paragraph,
        RtfRun run,
        string location,
        bool nonPrimary,
        int? paragraphBackground,
        OfficeContentSafetyBuilder builder,
        IDictionary<string, RtfContentSafetyTarget>? targets) {
        if (string.IsNullOrWhiteSpace(run.Text)) return;
        RtfEffectiveCharacterStyle effective = ResolveRtfCharacterStyle(document, paragraph, run);
        OfficeContentConcealmentKind? kind = null;
        string? evidence = null;
        if (run.Hidden) {
            kind = OfficeContentConcealmentKind.HiddenByProperty;
            evidence = "The effective RTF character state enables the native hidden-text control (\\v).";
        } else if (run.RevisionKind == RtfRevisionKind.Deleted) {
            kind = OfficeContentConcealmentKind.HiddenByProperty;
            evidence = "The RTF run is retained as deleted revision content rather than ordinary current text.";
        } else if (effective.FontSize.HasValue && effective.FontSize.Value <= builder.Options.MaximumTinyFontSizePoints) {
            kind = OfficeContentConcealmentKind.TinyText;
            evidence = "The effective RTF font size is " + effective.FontSize.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt.";
        } else if (run.CharacterScalePercent.HasValue && run.CharacterScalePercent.Value <= 1) {
            kind = OfficeContentConcealmentKind.ZeroDimension;
            evidence = "The RTF run uses a character scale of " + run.CharacterScalePercent.Value.ToString(CultureInfo.InvariantCulture) + " percent.";
        } else if (TryGetRtfContrast(document, effective.ForegroundColorIndex, run.HighlightColorIndex ?? run.CharacterBackgroundColorIndex ?? paragraphBackground, out double ratio, out string colors) &&
                   ratio < builder.Options.MinimumVisibleContrastRatio) {
            kind = OfficeContentConcealmentKind.LowContrastText;
            evidence = colors + " has contrast ratio " + ratio.ToString("0.###", CultureInfo.InvariantCulture) + ".";
        } else if (nonPrimary && builder.Options.IncludeNonPrimaryContent) {
            kind = OfficeContentConcealmentKind.NonPrimaryContent;
            evidence = "The RTF text is stored in a note, object fallback, or another non-primary story.";
        }

        if (kind.HasValue) {
            OfficeContentSafetyFinding finding = builder.Add(kind.Value, OfficeContentSafetyRisk.ContextDependent, location, evidence!, run.Text, OfficeContentCleanupCapability.RemoveText, inspectTextIntegrityEvidence: false);
            if (targets != null) targets[finding.Id] = RtfContentSafetyTarget.ForRun(run);
        }
        IReadOnlyList<OfficeContentSafetyFinding> unicode = kind.HasValue
            ? builder.InspectChargedTextIntegrity(location + "/Text", run.Text, OfficeContentCleanupCapability.RemoveText)
            : builder.InspectVisibleText(location + "/Text", run.Text, OfficeContentCleanupCapability.RemoveText);
        if (targets != null) foreach (OfficeContentSafetyFinding item in unicode) targets[item.Id] = RtfContentSafetyTarget.ForRunRange(run, item);
    }

    private static RtfEffectiveCharacterStyle ResolveRtfCharacterStyle(RtfDocument document, RtfParagraph paragraph, RtfRun run) {
        var effective = new RtfEffectiveCharacterStyle();
        ApplyRtfStyleChain(document, paragraph.StyleId, effective);
        ApplyRtfStyleChain(document, run.StyleId, effective);
        if (run.FontSize.HasValue) effective.FontSize = run.FontSize;
        if (run.ForegroundColorIndex.HasValue) effective.ForegroundColorIndex = run.ForegroundColorIndex;
        return effective;
    }

    private static int? ResolveRtfParagraphStyleBackground(RtfDocument document, int? styleId) {
        var chain = new Stack<RtfStyle>();
        var visited = new HashSet<int>();
        RtfStyle? current = styleId.HasValue ? document.Styles.FirstOrDefault(item => item.Id == styleId.Value) : null;
        while (current != null && visited.Add(current.Id)) {
            chain.Push(current);
            current = current.BasedOnStyleId.HasValue ? document.Styles.FirstOrDefault(item => item.Id == current.BasedOnStyleId.Value) : null;
        }
        int? background = null;
        while (chain.Count > 0) background = chain.Pop().BackgroundColorIndex ?? background;
        return background;
    }

    private static void ApplyRtfStyleChain(RtfDocument document, int? styleId, RtfEffectiveCharacterStyle target) {
        var chain = new Stack<RtfStyle>();
        var visited = new HashSet<int>();
        RtfStyle? current = styleId.HasValue ? document.Styles.FirstOrDefault(item => item.Id == styleId.Value) : null;
        while (current != null && visited.Add(current.Id)) {
            chain.Push(current);
            current = current.BasedOnStyleId.HasValue ? document.Styles.FirstOrDefault(item => item.Id == current.BasedOnStyleId.Value) : null;
        }
        while (chain.Count > 0) {
            RtfStyle style = chain.Pop();
            if (style.FontSize.HasValue) target.FontSize = style.FontSize;
            if (style.ForegroundColorIndex.HasValue) target.ForegroundColorIndex = style.ForegroundColorIndex;
        }
    }

    private static bool TryGetRtfContrast(RtfDocument document, int? foregroundIndex, int? backgroundIndex, out double ratio, out string evidence) {
        OfficeColor foreground = ResolveRtfColor(document, foregroundIndex, OfficeColor.Black);
        OfficeColor background = ResolveRtfColor(document, backgroundIndex, OfficeColor.White);
        ratio = OfficeColorContrast.ContrastRatio(foreground, background);
        evidence = "Effective RTF foreground #" + foreground.ToRgbHex() + " against background #" + background.ToRgbHex();
        return foregroundIndex.HasValue || backgroundIndex.HasValue;
    }

    private static OfficeColor ResolveRtfColor(RtfDocument document, int? index, OfficeColor fallback) {
        if (!index.HasValue || index.Value <= 0 || index.Value > document.Colors.Count) return fallback;
        RtfColor color = document.Colors[index.Value - 1];
        return OfficeColor.FromRgb(color.Red, color.Green, color.Blue);
    }

    private sealed class RtfEffectiveCharacterStyle {
        internal double? FontSize { get; set; }
        internal int? ForegroundColorIndex { get; set; }
    }

    private sealed class RtfContentSafetyTarget {
        private readonly RtfRun? _run;
        private readonly RtfDocument? _document;
        private readonly int? _offset;
        private readonly int? _length;
        private readonly string? _expected;
        private RtfContentSafetyTarget(RtfRun? run, RtfDocument? document, int? offset = null, int? length = null, string? expected = null) { _run = run; _document = document; _offset = offset; _length = length; _expected = expected; }
        internal static RtfContentSafetyTarget ForRun(RtfRun run) => new RtfContentSafetyTarget(run, null);
        internal static RtfContentSafetyTarget ForHtmlEncapsulation(RtfDocument document) => new RtfContentSafetyTarget(null, document);
        internal static RtfContentSafetyTarget ForHtmlEncapsulationRange(RtfDocument document, OfficeContentSafetyFinding finding) => new RtfContentSafetyTarget(
            null, document, finding.SourceTextOffset, finding.SourceTextLength,
            document.HtmlEncapsulation!.Html.Substring(finding.SourceTextOffset!.Value, finding.SourceTextLength!.Value));
        internal static RtfContentSafetyTarget ForRunRange(RtfRun run, OfficeContentSafetyFinding finding) => new RtfContentSafetyTarget(
            run, null, finding.SourceTextOffset, finding.SourceTextLength,
            run.Text.Substring(finding.SourceTextOffset!.Value, finding.SourceTextLength!.Value));
        internal void Remove() {
            if (_run != null && _offset.HasValue && _length.HasValue) {
                if (_offset.Value > _run.Text.Length - _length.Value || !string.Equals(_run.Text.Substring(_offset.Value, _length.Value), _expected, StringComparison.Ordinal)) {
                    throw new InvalidOperationException("The selected Unicode text range no longer matches the inspected RTF run.");
                }
                _run.Text = _run.Text.Remove(_offset.Value, _length.Value);
            } else if (_run != null) _run.Text = string.Empty;
            else if (_document?.HtmlEncapsulation != null && _offset.HasValue && _length.HasValue) {
                RtfHtmlEncapsulation current = _document.HtmlEncapsulation;
                if (_offset.Value > current.Html.Length - _length.Value || !string.Equals(current.Html.Substring(_offset.Value, _length.Value), _expected, StringComparison.Ordinal)) {
                    throw new InvalidOperationException("The selected Unicode text range no longer matches the inspected RTF HTML encapsulation.");
                }
                _document.HtmlEncapsulation = new RtfHtmlEncapsulation(current.Version, current.Html.Remove(_offset.Value, _length.Value));
            } else if (_document != null) _document.HtmlEncapsulation = null;
        }
    }
}
