using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Word;

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
        return SelectAdapter(data, effective).Detection;
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
        (ILegacyWordAdapter adapter, LegacyWordDetection detection) = SelectAdapter(data, effective);
        cancellationToken.ThrowIfCancellationRequested();
        LegacyWordModel model = adapter.Parse(data, effective.Limits, cancellationToken);
        if (effective.RequireStructured && model.Quality != OfficeLegacyImportQuality.Structured) {
            throw new InvalidDataException($"The {detection.ProfileId} adapter produced salvage quality while structured import was required.");
        }

        WordDocument document = Project(model, cancellationToken);
        string text = string.Join(Environment.NewLine, model.Paragraphs.Select(static paragraph => paragraph.Text)
            .Concat(model.Notes.Select(static note => "[" + note.Kind + "] " + note.Text)));
        var report = new OfficeLegacyImportReport(detection.ProfileId, model.Quality, model.Findings, model.InertContent, model.Paragraphs.Count + model.Notes.Count + model.Resources.Count);
        return new LegacyWordImportResult(document, detection, report, text,
            new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(model.Metadata, StringComparer.OrdinalIgnoreCase)),
            new LegacyWordContent(model));
    }

    private static WordDocument Project(LegacyWordModel model, CancellationToken cancellationToken) {
        WordDocument document = WordDocument.Create();
        try {
            WordList? activeList = null;
            foreach (LegacyWordParagraph source in model.Paragraphs) {
                cancellationToken.ThrowIfCancellationRequested();
                if (source.IsList) {
                    activeList ??= document.AddListBulleted();
                    WordParagraph paragraph = activeList.AddItem(string.Empty, Math.Max(0, Math.Min(8, source.ListLevel)));
                    ProjectParagraph(source, paragraph);
                } else {
                    activeList = null;
                    WordParagraph paragraph = document.AddParagraph();
                    ProjectParagraph(source, paragraph);
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

    private static void ProjectParagraph(LegacyWordParagraph source, WordParagraph paragraph) {
        foreach (LegacyWordRun sourceRun in source.Runs) {
            WordParagraph run = paragraph.AddFormattedText(sourceRun.Text, sourceRun.Bold, sourceRun.Italic, sourceRun.Underline);
            if (sourceRun.Strike) run.SetStrike();
            if (sourceRun.VerticalPosition.HasValue) run.SetVerticalTextAlignment(sourceRun.VerticalPosition);
            if (sourceRun.FontSizePoints.HasValue) run.SetFontSize(sourceRun.FontSizePoints.Value);
            if (!string.IsNullOrWhiteSpace(sourceRun.FontFamily)) run.SetFontFamily(sourceRun.FontFamily!);
            if (!string.IsNullOrWhiteSpace(sourceRun.ColorHex)) run.SetColorHex(sourceRun.ColorHex!);
        }
        if (source.Alignment.HasValue) paragraph.SetAlignment(source.Alignment.Value);
        paragraph.PageBreakBefore = source.PageBreakBefore;
        paragraph.KeepWithNext = source.KeepWithNext;
        paragraph.KeepLinesTogether = source.KeepLinesTogether;
        paragraph.LineSpacingPoints = source.LineSpacingPoints;
        paragraph.LineSpacingBeforePoints = source.SpacingBeforePoints;
        paragraph.LineSpacingAfterPoints = source.SpacingAfterPoints;
    }

    private static (ILegacyWordAdapter Adapter, LegacyWordDetection Detection) SelectAdapter(byte[] data, LegacyWordImportOptions options) {
        if (options.FormatHint.HasValue) {
            ILegacyWordAdapter hinted = Adapters.Single(adapter => adapter.Format == options.FormatHint.Value);
            int confidence = hinted.Probe(data, options.SourceName, out string evidence);
            return (hinted, new LegacyWordDetection(hinted.Format, hinted.GetProfileId(data), Math.Max(1, confidence),
                confidence == 0 ? "Explicit caller format hint." : evidence + " Explicit caller format hint confirmed the family."));
        }

        ILegacyWordAdapter? selected = null;
        string selectedReason = string.Empty;
        int selectedConfidence = 0;
        foreach (ILegacyWordAdapter adapter in Adapters) {
            int confidence = adapter.Probe(data, options.SourceName, out string reason);
            if (confidence > selectedConfidence) {
                selected = adapter;
                selectedConfidence = confidence;
                selectedReason = reason;
            }
        }
        if (selected == null || selectedConfidence < 50) {
            throw new InvalidDataException("The source does not match a supported bounded legacy-word profile. Supply FormatHint only when the family is known.");
        }
        return (selected, new LegacyWordDetection(selected.Format, selected.GetProfileId(data), selectedConfidence, selectedReason));
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
