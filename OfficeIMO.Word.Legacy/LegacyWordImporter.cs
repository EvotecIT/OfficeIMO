using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OpenXmlStyle = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace OfficeIMO.Word.Legacy;

/// <summary>Detects and imports selected legacy word-processing formats without executing source content.</summary>
public static class LegacyWordImporter {
    private static readonly ILegacyWordAdapter[] Adapters = {
        new WordPerfectAdapter(),
        new WordStarAdapter(),
        new AmiProAdapter(),
        new LotusWordProAdapter(),
        new MicrosoftWorksWordAdapter(),
        new MicrosoftWriteAdapter(),
        new WordForDosAdapter()
    };

    /// <summary>Detects a legacy-word source from a file.</summary>
    public static LegacyWordDetection Detect(string path, LegacyWordImportOptions? options = null, CancellationToken cancellationToken = default) {
        LegacyWordImportOptions effective = Prepare(options, path);
        byte[] data = OfficeLegacyImportBuffer.ReadAll(path, effective.Limits, cancellationToken);
        return Detect(data, effective, cancellationToken);
    }

    /// <summary>Detects a legacy-word source from bytes.</summary>
    public static LegacyWordDetection Detect(byte[] data, LegacyWordImportOptions? options = null, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        LegacyWordImportOptions effective = Prepare(options, options?.SourceName);
        if (data.Length > effective.Limits.MaxInputBytes) throw new InvalidDataException("Legacy word input exceeds the configured byte limit.");
        cancellationToken.ThrowIfCancellationRequested();
        return SelectAdapter(data, effective, cancellationToken).Detection;
    }

    /// <summary>Imports a legacy-word file into a normal editable <see cref="WordDocument"/>.</summary>
    public static LegacyWordImportResult Import(string path, LegacyWordImportOptions? options = null, CancellationToken cancellationToken = default) {
        LegacyWordImportOptions effective = Prepare(options, path);
        byte[] data = OfficeLegacyImportBuffer.ReadAll(path, effective.Limits, cancellationToken);
        return Import(data, effective, cancellationToken);
    }

    /// <summary>Imports a legacy-word stream into a normal editable <see cref="WordDocument"/>.</summary>
    public static LegacyWordImportResult Import(Stream stream, LegacyWordImportOptions? options = null, CancellationToken cancellationToken = default) {
        LegacyWordImportOptions effective = Prepare(options, options?.SourceName);
        byte[] data = OfficeLegacyImportBuffer.ReadAll(stream, effective.Limits, cancellationToken);
        return Import(data, effective, cancellationToken);
    }

    /// <summary>Imports legacy-word bytes into a normal editable <see cref="WordDocument"/>.</summary>
    public static LegacyWordImportResult Import(byte[] data, LegacyWordImportOptions? options = null, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        LegacyWordImportOptions effective = Prepare(options, options?.SourceName);
        if (data.Length > effective.Limits.MaxInputBytes) throw new InvalidDataException("Legacy word input exceeds the configured byte limit.");
        (ILegacyWordAdapter adapter, LegacyWordDetection detection) = SelectAdapter(data, effective, cancellationToken);
        LegacyWordModel model = adapter.Parse(data, effective.Limits, cancellationToken);
        if (effective.RequireStructured && model.Quality != OfficeLegacyImportQuality.Structured) {
            throw new InvalidDataException($"The {detection.ProfileId} adapter produced salvage quality while structured import was required.");
        }

        string text = BuildPlainText(model, effective.Limits, cancellationToken);
        var report = new OfficeLegacyImportReport(detection.ProfileId, model.Quality, model.Findings, model.InertContent, model.Paragraphs.Count + model.Styles.Count + model.Notes.Count + model.Resources.Count);
        var metadata = new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(model.Metadata, StringComparer.OrdinalIgnoreCase));
        var content = new LegacyWordContent(model);
        WordDocument? document = null;
        try {
            document = Project(model, cancellationToken);
            return new LegacyWordImportResult(document, detection, report, text, metadata, content);
        } catch {
            document?.Dispose();
            throw;
        }
    }

    private static string BuildPlainText(LegacyWordModel model, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) {
        var text = new StringBuilder(Math.Min(limits.MaxTextCharacters, 4096));
        bool hasEntry = false;
        foreach (LegacyWordParagraph paragraph in model.Paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            AppendPlainTextEntry(text, paragraph.Text, ref hasEntry, limits.MaxTextCharacters);
        }
        foreach (LegacyWordNote note in model.Notes) {
            cancellationToken.ThrowIfCancellationRequested();
            AppendPlainTextEntry(text, "[" + note.Kind + "] " + note.Text, ref hasEntry, limits.MaxTextCharacters);
        }
        return text.ToString();
    }

    private static void AppendPlainTextEntry(StringBuilder target, string value, ref bool hasEntry, int maxCharacters) {
        int separatorLength = hasEntry ? 1 : 0;
        if (value.Length > maxCharacters - target.Length - separatorLength) {
            throw new InvalidDataException("Legacy word plain-text projection exceeds the configured character limit.");
        }
        if (hasEntry) target.Append('\n');
        target.Append(value);
        hasEntry = true;
    }

    private static WordDocument Project(LegacyWordModel model, CancellationToken cancellationToken) {
        WordDocument document = WordDocument.Create();
        try {
            document.EnsureStyleDefinitionsInitialized();
            WordList? activeList = null;
            var styleIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var usedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Styles? existingStyles = document.OpenXmlDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (existingStyles != null) {
                foreach (OpenXmlStyle style in existingStyles.Elements<OpenXmlStyle>()) {
                    string? styleId = style.StyleId?.Value;
                    if (!string.IsNullOrWhiteSpace(styleId)) usedStyleIds.Add(styleId!);
                }
            }
            var recoveredStyles = new Dictionary<string, LegacyWordStyle>(StringComparer.OrdinalIgnoreCase);
            foreach (LegacyWordStyle sourceStyle in model.Styles) {
                cancellationToken.ThrowIfCancellationRequested();
                string styleId = GetOrCreateLegacyStyleId(sourceStyle.Name, styleIds, usedStyleIds);
                recoveredStyles[sourceStyle.Name] = sourceStyle;
                EnsureLegacyParagraphStyle(document, styleId, sourceStyle.Name, sourceStyle);
            }
            foreach (LegacyWordParagraph source in model.Paragraphs) {
                cancellationToken.ThrowIfCancellationRequested();
                if (source.IsList) {
                    activeList ??= document.AddListBulleted();
                    WordParagraph paragraph = activeList.AddItem(string.Empty, Math.Max(0, Math.Min(8, source.ListLevel)));
                    ProjectParagraph(source, paragraph, document, styleIds, usedStyleIds, recoveredStyles, cancellationToken);
                } else {
                    activeList = null;
                    WordParagraph paragraph = document.AddParagraph();
                    ProjectParagraph(source, paragraph, document, styleIds, usedStyleIds, recoveredStyles, cancellationToken);
                }
            }
            foreach (LegacyWordNote note in model.Notes) {
                cancellationToken.ThrowIfCancellationRequested();
                activeList = null;
                WordParagraph paragraph = document.AddParagraph("[Recovered " + note.Kind + "] " + note.Text);
                paragraph.AddComment("Legacy source", "LS", note.Kind + " recovered without its original source anchor.");
            }
            if (model.Paragraphs.Count == 0) document.AddParagraph(string.Empty);
            return document;
        } catch {
            document.Dispose();
            throw;
        }
    }

    private static void ProjectParagraph(LegacyWordParagraph source, WordParagraph paragraph, WordDocument document,
        IDictionary<string, string> styleIds, ISet<string> usedStyleIds, IReadOnlyDictionary<string, LegacyWordStyle> recoveredStyles,
        CancellationToken cancellationToken) {
        LegacyWordStyle? recoveredStyle = null;
        if (!string.IsNullOrWhiteSpace(source.StyleName)) {
            recoveredStyles.TryGetValue(source.StyleName!, out recoveredStyle);
        }
        int runIndex = 0;
        foreach (LegacyWordRun sourceRun in source.Runs) {
            if ((runIndex++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            ProjectRun(sourceRun, paragraph, recoveredStyle, document, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Alignment.HasValue) paragraph.SetAlignment(source.Alignment.Value);
        paragraph.PageBreakBefore = source.PageBreakBefore;
        paragraph.KeepWithNext = source.KeepWithNext;
        paragraph.KeepLinesTogether = source.KeepLinesTogether;
        paragraph.LineSpacingPoints = source.LineSpacingPoints;
        paragraph.LineSpacingBeforePoints = source.SpacingBeforePoints;
        paragraph.LineSpacingAfterPoints = source.SpacingAfterPoints;
        if (!string.IsNullOrWhiteSpace(source.StyleName)) {
            string styleId = GetOrCreateLegacyStyleId(source.StyleName!, styleIds, usedStyleIds);
            paragraph.SetStyleId(styleId);
            EnsureLegacyParagraphStyle(document, styleId, source.StyleName!, recoveredStyle);
        }
    }

    private static void ProjectRun(
        LegacyWordRun sourceRun,
        WordParagraph paragraph,
        LegacyWordStyle? recoveredStyle,
        WordDocument document,
        CancellationToken cancellationToken) {
        string normalized = sourceRun.Text.Replace("\r\n", "\n").Replace('\r', '\n');
        for (int segmentStart = 0; segmentStart <= normalized.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int segmentEnd = normalized.IndexOf('\n', segmentStart);
            if (segmentEnd < 0) segmentEnd = normalized.Length;
            if (segmentStart > 0) paragraph.AddBreak();
            int segmentLength = segmentEnd - segmentStart;
            if (segmentLength > 0 || normalized.Length == 0) {
                WordParagraph run = paragraph.AddFormattedText(
                    normalized.Substring(segmentStart, segmentLength),
                    sourceRun.Bold, sourceRun.Italic, sourceRun.Underline);
                if (sourceRun.Strike) run.SetStrike();
                if (sourceRun.VerticalPosition.HasValue) run.SetVerticalTextAlignment(sourceRun.VerticalPosition);
                if (sourceRun.FontSizePoints.HasValue) run.FontSizePoints = sourceRun.FontSizePoints.Value;
                if (!string.IsNullOrWhiteSpace(sourceRun.FontFamily)) run.SetFontFamily(sourceRun.FontFamily!);
                if (!string.IsNullOrWhiteSpace(sourceRun.ColorHex)) run.SetColorHex(sourceRun.ColorHex!);
                ApplyRecoveredStyleRunOverrides(run, sourceRun, recoveredStyle, document);
            }
            if (segmentEnd == normalized.Length) break;
            segmentStart = segmentEnd + 1;
        }
    }

    private static void ApplyRecoveredStyleRunOverrides(
        WordParagraph run,
        LegacyWordRun sourceRun,
        LegacyWordStyle? recoveredStyle,
        WordDocument document) {
        if (recoveredStyle == null) return;
        bool resetBold = recoveredStyle.Bold && !sourceRun.Bold;
        bool resetItalic = recoveredStyle.Italic && !sourceRun.Italic;
        bool resetUnderline = recoveredStyle.Underline.HasValue && !sourceRun.Underline.HasValue;
        bool resetFontFamily = recoveredStyle.FontFamily != null && sourceRun.FontFamily == null;
        bool resetFontSize = recoveredStyle.FontSizePoints.HasValue && !sourceRun.FontSizePoints.HasValue;
        bool resetColor = recoveredStyle.ColorHex != null && sourceRun.ColorHex == null;
        if (!resetBold && !resetItalic && !resetUnderline && !resetFontFamily && !resetFontSize && !resetColor) return;

        RunProperties properties = run._runProperties ??= new RunProperties();
        if (resetBold) {
            properties.Bold = new Bold { Val = false };
            properties.BoldComplexScript = new BoldComplexScript { Val = false };
        }
        if (resetItalic) {
            properties.Italic = new Italic { Val = false };
            properties.ItalicComplexScript = new ItalicComplexScript { Val = false };
        }
        if (resetUnderline) {
            properties.Underline = new Underline { Val = UnderlineValues.None };
        }

        RunPropertiesBaseStyle? defaults = document.OpenXmlDocument.MainDocumentPart?
            .StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (resetFontFamily) {
            properties.RunFonts = defaults?.RunFonts != null
                ? (RunFonts)defaults.RunFonts.CloneNode(true)
                : new RunFonts {
                    AsciiTheme = ThemeFontValues.MinorHighAnsi,
                    HighAnsiTheme = ThemeFontValues.MinorHighAnsi,
                    EastAsiaTheme = ThemeFontValues.MinorHighAnsi,
                    ComplexScriptTheme = ThemeFontValues.MinorBidi
                };
        }
        if (resetFontSize) {
            properties.FontSize = defaults?.FontSize != null
                ? (FontSize)defaults.FontSize.CloneNode(true)
                : new FontSize { Val = "22" };
            properties.FontSizeComplexScript = defaults?.FontSizeComplexScript != null
                ? (FontSizeComplexScript)defaults.FontSizeComplexScript.CloneNode(true)
                : new FontSizeComplexScript { Val = "22" };
        }
        if (resetColor) {
            properties.Color = new Color { Val = "auto" };
        }
    }

    private static string GetOrCreateLegacyStyleId(string styleName, IDictionary<string, string> styleIds, ISet<string> usedStyleIds) {
        if (styleIds.TryGetValue(styleName, out string? existing)) return existing;
        var identifier = new StringBuilder("Legacy");
        foreach (char value in styleName) {
            if (identifier.Length >= 56) break;
            if ((value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z') || (value >= '0' && value <= '9')) identifier.Append(value);
        }
        if (identifier.Length == 6) identifier.Append("Style");
        string basis = identifier.ToString();
        string candidate = basis;
        for (int suffix = 2; !usedStyleIds.Add(candidate); suffix++) candidate = basis + suffix.ToString(CultureInfo.InvariantCulture);
        styleIds.Add(styleName, candidate);
        return candidate;
    }

    private static void EnsureLegacyParagraphStyle(WordDocument document, string styleId, string styleName, LegacyWordStyle? recoveredStyle) {
        MainDocumentPart mainPart = document.OpenXmlDocument.MainDocumentPart
            ?? throw new InvalidDataException("The projected Word document has no main document part.");
        StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart
            ?? throw new InvalidDataException("The projected Word document has no style definitions part.");
        Styles styles = stylePart.Styles ??= new Styles();
        if (styles.Elements<OpenXmlStyle>().Any(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase))) return;
        var projected = new OpenXmlStyle(
            new StyleName { Val = styleName },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" }) {
            Type = StyleValues.Paragraph,
            StyleId = styleId,
            CustomStyle = true
        };
        if (recoveredStyle != null) {
            var paragraphProperties = new StyleParagraphProperties();
            if (recoveredStyle.Alignment.HasValue) paragraphProperties.Justification = new Justification { Val = ToOpenXmlAlignment(recoveredStyle.Alignment.Value) };
            if (recoveredStyle.LineSpacingPoints.HasValue || recoveredStyle.SpacingBeforePoints.HasValue || recoveredStyle.SpacingAfterPoints.HasValue) {
                paragraphProperties.SpacingBetweenLines = new SpacingBetweenLines {
                    Line = ToTwips(recoveredStyle.LineSpacingPoints),
                    Before = ToTwips(recoveredStyle.SpacingBeforePoints),
                    After = ToTwips(recoveredStyle.SpacingAfterPoints)
                };
            }
            if (recoveredStyle.PageBreakBefore) paragraphProperties.PageBreakBefore = new PageBreakBefore();
            if (recoveredStyle.KeepWithNext) paragraphProperties.KeepNext = new KeepNext();
            if (recoveredStyle.KeepLinesTogether) paragraphProperties.KeepLines = new KeepLines();
            if (paragraphProperties.HasChildren) projected.Append(paragraphProperties);

            var runProperties = new StyleRunProperties {
                Bold = new Bold { Val = recoveredStyle.Bold },
                Italic = new Italic { Val = recoveredStyle.Italic }
            };
            if (!string.IsNullOrWhiteSpace(recoveredStyle.FontFamily)) {
                runProperties.RunFonts = new RunFonts {
                    Ascii = recoveredStyle.FontFamily,
                    HighAnsi = recoveredStyle.FontFamily,
                    EastAsia = recoveredStyle.FontFamily,
                    ComplexScript = recoveredStyle.FontFamily
                };
            }
            if (recoveredStyle.FontSizePoints.HasValue) {
                string halfPoints = Math.Round(recoveredStyle.FontSizePoints.Value * 2d, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
                runProperties.FontSize = new FontSize { Val = halfPoints };
                runProperties.FontSizeComplexScript = new FontSizeComplexScript { Val = halfPoints };
            }
            if (!string.IsNullOrWhiteSpace(recoveredStyle.ColorHex)) runProperties.Color = new Color { Val = recoveredStyle.ColorHex };
            if (recoveredStyle.Underline.HasValue) runProperties.Underline = new Underline { Val = ToOpenXmlUnderline(recoveredStyle.Underline.Value) };
            projected.Append(runProperties);
        }
        styles.Append(projected);
    }

    private static string? ToTwips(double? points) => points.HasValue
        ? Math.Round(points.Value * 20d, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture)
        : null;

    private static UnderlineValues ToOpenXmlUnderline(WordUnderlineStyle value) => value switch {
        WordUnderlineStyle.Double => UnderlineValues.Double,
        WordUnderlineStyle.Words => UnderlineValues.Words,
        WordUnderlineStyle.Single => UnderlineValues.Single,
        _ => throw new InvalidDataException($"Recovered underline style {value} is outside the DOCX projection profile.")
    };

    private static JustificationValues ToOpenXmlAlignment(WordParagraphAlignment value) => value switch {
        WordParagraphAlignment.Center => JustificationValues.Center,
        WordParagraphAlignment.Right => JustificationValues.Right,
        WordParagraphAlignment.Both => JustificationValues.Both,
        WordParagraphAlignment.Left => JustificationValues.Left,
        _ => throw new InvalidDataException($"Recovered paragraph alignment {value} is outside the DOCX projection profile.")
    };

    private static (ILegacyWordAdapter Adapter, LegacyWordDetection Detection) SelectAdapter(byte[] data, LegacyWordImportOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.FormatHint.HasValue) {
            ILegacyWordAdapter hinted = Adapters.Single(adapter => adapter.Format == options.FormatHint.Value);
            int confidence = hinted.Probe(data, options.SourceName, options.Limits, cancellationToken, out string evidence);
            return (hinted, new LegacyWordDetection(hinted.Format, hinted.GetProfileId(data, options.Limits, cancellationToken), Math.Max(1, confidence),
                confidence == 0 ? "Explicit caller format hint." : evidence + " Explicit caller format hint confirmed the family."));
        }

        ILegacyWordAdapter? selected = null;
        string selectedReason = string.Empty;
        int selectedConfidence = 0;
        foreach (ILegacyWordAdapter adapter in Adapters) {
            cancellationToken.ThrowIfCancellationRequested();
            int confidence = adapter.Probe(data, options.SourceName, options.Limits, cancellationToken, out string reason);
            if (confidence > selectedConfidence) {
                selected = adapter;
                selectedConfidence = confidence;
                selectedReason = reason;
            }
        }
        if (selected == null || selectedConfidence < 50) {
            throw new InvalidDataException("The source does not match a supported bounded legacy-word profile. Supply FormatHint only when the family is known.");
        }
        return (selected, new LegacyWordDetection(selected.Format, selected.GetProfileId(data, options.Limits, cancellationToken), selectedConfidence, selectedReason));
    }

    private static LegacyWordImportOptions Prepare(LegacyWordImportOptions? source, string? fallbackName) {
        var options = new LegacyWordImportOptions {
            Limits = (source?.Limits ?? new OfficeLegacyImportLimits()).Clone(),
            FormatHint = source?.FormatHint,
            SourceName = string.IsNullOrWhiteSpace(source?.SourceName) ? fallbackName : source!.SourceName,
            RequireStructured = source?.RequireStructured ?? false
        };
        options.Limits.Validate();
        return options;
    }
}
