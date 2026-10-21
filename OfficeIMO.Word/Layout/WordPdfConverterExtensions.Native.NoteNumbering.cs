using System.Collections.Generic;
using OfficeIMO.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Keeps reference identity distinct from labels that can repeat between note kinds or sections.</summary>
    private enum NativeNoteKind { Footnote, Endnote }
    private readonly record struct NativeNoteKey(NativeNoteKind Kind, long Id);

    private sealed class NativeNoteNumbering {
        private readonly Dictionary<NativeNoteKey, int> tokensById = new();
        private readonly Dictionary<int, string> labelsByToken = new();
        private readonly WordDocument document;
        private readonly W.Settings? settings;
        private readonly WordToPdfOptions? options;
        private int? nextFootnote;
        private int? nextEndnote;
        private bool hasFootnotes;
        private bool hasEndnotes;
        private W.NumberFormatValues footnoteFormat;
        private W.NumberFormatValues endnoteFormat;
        private bool footnotePageRestart;
        private bool endnotePageRestart;
        private bool pageRestartReported;
        private bool placementFallbackReported;
        private bool separatorMarkOmissionReported;
        private bool separatorDecorationOmissionReported;
        private WordSection? currentSection;
        private readonly Dictionary<NativeNoteKey, NativePageFootnoteSource> pageFootnotes = new();
        internal bool EndnotesAtDocumentEnd { get; }
        internal bool TryGetPageFootnote(NativeNoteKey key, out OfficeIMO.Pdf.PdfFootnoteDefinition? definition) =>
            TryGetPageFootnoteDefinition(key, out definition);
        private bool TryGetPageFootnoteDefinition(NativeNoteKey key, out OfficeIMO.Pdf.PdfFootnoteDefinition? definition) {
            bool found = pageFootnotes.TryGetValue(key, out NativePageFootnoteSource? source);
            definition = found && ReferenceEquals(source!.Section, currentSection) ? source.Definition : null;
            return definition != null;
        }
        internal void SetPageFootnote(NativePageFootnoteSource source) => pageFootnotes[source.Key] = source;

        internal NativeNoteNumbering(WordDocument document, WordToPdfOptions? options) {
            this.document = document;
            settings = document._wordprocessingDocument?.MainDocumentPart?.DocumentSettingsPart?.Settings;
            this.options = options;
            EndnotesAtDocumentEnd = document.EndnoteSettings.Position == WordEndnotePosition.DocumentEnd;
        }

        internal NativeNoteNumbering CreateRunningReferenceSnapshot(WordSection section) {
            var snapshot = new NativeNoteNumbering(document, null);
            snapshot.BeginSection(section);
            foreach (var pair in tokensById) snapshot.tokensById.Add(pair.Key, pair.Value);
            foreach (var pair in labelsByToken) snapshot.labelsByToken.Add(pair.Key, pair.Value);
            return snapshot;
        }

        internal void BeginSection(WordSection section) {
            currentSection = section;
            W.FootnoteProperties? foot = section._sectionProperties.GetFirstChild<W.FootnoteProperties>();
            W.EndnoteProperties? end = section._sectionProperties.GetFirstChild<W.EndnoteProperties>();
            W.FootnoteDocumentWideProperties? documentFoot = settings?.GetFirstChild<W.FootnoteDocumentWideProperties>();
            W.EndnoteDocumentWideProperties? documentEnd = settings?.GetFirstChild<W.EndnoteDocumentWideProperties>();
            footnoteFormat = foot?.NumberingFormat?.Val?.Value ?? documentFoot?.NumberingFormat?.Val?.Value ?? W.NumberFormatValues.Decimal;
            endnoteFormat = end?.NumberingFormat?.Val?.Value ?? documentEnd?.NumberingFormat?.Val?.Value ?? W.NumberFormatValues.LowerRoman;
            W.RestartNumberValues? footRestart = foot?.NumberingRestart?.Val?.Value ?? documentFoot?.NumberingRestart?.Val?.Value;
            W.RestartNumberValues? endRestart = end?.NumberingRestart?.Val?.Value ?? documentEnd?.NumberingRestart?.Val?.Value;
            if (!hasFootnotes || footRestart == W.RestartNumberValues.EachSection)
                nextFootnote = foot?.NumberingStart?.Val?.Value ?? documentFoot?.NumberingStart?.Val?.Value ?? 1;
            if (!hasEndnotes || endRestart == W.RestartNumberValues.EachSection)
                nextEndnote = end?.NumberingStart?.Val?.Value ?? documentEnd?.NumberingStart?.Val?.Value ?? 1;
            footnotePageRestart = footRestart == W.RestartNumberValues.EachPage;
            endnotePageRestart = endRestart == W.RestartNumberValues.EachPage;
        }

        internal void ReportSeparatorMarkOmission() {
            if (separatorMarkOmissionReported || options == null) return;
            separatorMarkOmissionReported = true;
            AddNativeExportWarning(options, "NativeFootnoteSeparatorMarkOmitted", "footnote separator",
                "The separator paragraph and its spacing are retained, but its rule mark is not painted because the selected font does not provide usable strike-band metrics.");
        }

        internal void ReportSeparatorDecorationOmission(string reason) {
            if (separatorDecorationOmissionReported || options == null) return;
            separatorDecorationOmissionReported = true;
            AddNativeExportWarning(options, "NativeFootnoteSeparatorDecorationOmitted", "footnote separator",
                "The separator rule is retained, but an additional authored decoration is not painted. " + reason);
        }

        internal void ReportPlacementFallback(string? reason = null) {
            if (placementFallbackReported || options == null) return;
            placementFallbackReported = true;
            AddNativeExportWarning(options, "NativeFootnotePlacementApproximated", "footnotes",
                "This footnote profile uses plain text in document flow after its section. Page placement and original note formatting are approximated." +
                (reason == null ? string.Empty : " " + reason));
        }

        internal bool ContainsKey(NativeNoteKey key) => tokensById.ContainsKey(key);
        internal bool TryGetValue(NativeNoteKey key, out int token) => tokensById.TryGetValue(key, out token);
        internal string GetLabel(int token) => labelsByToken[token];

        internal string Add(NativeNoteKey key, bool endnote) {
            int number = (endnote ? nextEndnote : nextFootnote) ?? 1;
            if (endnote) { nextEndnote = checked(number + 1); hasEndnotes = true; }
            else { nextFootnote = checked(number + 1); hasFootnotes = true; }
            W.NumberFormatValues format = endnote ? endnoteFormat : footnoteFormat;
            OfficeNumberStyle style;
            if (format == W.NumberFormatValues.LowerRoman) style = OfficeNumberStyle.LowerRoman;
            else if (format == W.NumberFormatValues.UpperRoman) style = OfficeNumberStyle.UpperRoman;
            else if (format == W.NumberFormatValues.LowerLetter) style = OfficeNumberStyle.RepeatedLowerLetter;
            else if (format == W.NumberFormatValues.UpperLetter) style = OfficeNumberStyle.RepeatedUpperLetter;
            else {
                style = OfficeNumberStyle.Decimal;
                if (format != W.NumberFormatValues.Decimal && options != null)
                    AddNativeExportWarning(options, "NativeNoteNumberFormatUnsupported", endnote ? "endnote" : "footnote",
                        "The note numbering format is outside the supported decimal, Roman and letter formats; a decimal label was used.");
            }
            if ((endnote ? endnotePageRestart : footnotePageRestart) && !pageRestartReported && options != null) {
                AddNativeExportWarning(options, "NativeNotePageRestartApproximated", "notes",
                    "Page-based note numbering restart is not applied; note labels continue through the section.");
                pageRestartReported = true;
            }
            // Word permits a zero DOCX note start. Decimal displays 0; Roman and
            // letter formats leave that marker blank before continuing at 1.
            // Keep the shared formatter's positive page-number contract intact.
            string label = number == 0 ? (style == OfficeNumberStyle.Decimal ? "0" : string.Empty) :
                OfficeNumberFormatter.Format(number, style);
            int token = tokensById.Count + 1;
            tokensById.Add(key, token);
            labelsByToken.Add(token, label);
            return label;
        }
    }
}
