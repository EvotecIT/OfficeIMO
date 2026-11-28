using OfficeIMO.Core;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Keeps reference identity distinct from labels that can repeat between note kinds or sections.</summary>
    private enum NativeNoteKind { Footnote, Endnote }
    private readonly record struct NativeNoteKey(NativeNoteKind Kind, long Id);

    private sealed class NativeNoteNumbering {
        private const int MaximumGeneratedLabelCharacters = 100_000;
        private readonly Dictionary<NativeNoteKey, int> tokensById = new();
        private readonly Dictionary<int, string> labelsByToken = new();
        private int generatedLabelCharacters;
        private readonly WordDocument document;
        private readonly OfficeXmlNode settings;
        private readonly WordPdfExportOptions? options;
        private int? nextFootnote;
        private int? nextEndnote;
        private bool hasFootnotes;
        private bool hasEndnotes;
        private string footnoteFormat = "decimal";
        private string endnoteFormat = "lowerRoman";
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

        internal NativeNoteNumbering(WordDocument document, WordPdfExportOptions? options) {
            this.document = document;
            settings = document.SettingsNode();
            this.options = options;
            EndnotesAtDocumentEnd = document.EndnoteSettings.Position == WordEndnotePosition.DocumentEnd;
        }

        internal NativeNoteNumbering CreateRunningReferenceSnapshot(WordSection section) {
            var snapshot = new NativeNoteNumbering(document, null);
            snapshot.BeginSection(section);
            foreach (var pair in tokensById) snapshot.tokensById.Add(pair.Key, pair.Value);
            foreach (var pair in labelsByToken) snapshot.labelsByToken.Add(pair.Key, pair.Value);
            snapshot.generatedLabelCharacters = generatedLabelCharacters;
            return snapshot;
        }

        internal void BeginSection(WordSection section) {
            currentSection = section;
            OfficeXmlNode foot = section.Node().WChild("footnotePr");
            OfficeXmlNode end = section.Node().WChild("endnotePr");
            OfficeXmlNode documentFoot = settings.WChild("footnotePr");
            OfficeXmlNode documentEnd = settings.WChild("endnotePr");
            footnoteFormat = foot.ChildVal("numFmt") ?? documentFoot.ChildVal("numFmt") ?? "decimal";
            endnoteFormat = end.ChildVal("numFmt") ?? documentEnd.ChildVal("numFmt") ?? "lowerRoman";
            string? footRestart = foot.ChildVal("numRestart") ?? documentFoot.ChildVal("numRestart");
            string? endRestart = end.ChildVal("numRestart") ?? documentEnd.ChildVal("numRestart");
            if (!hasFootnotes || footRestart == "eachSect")
                nextFootnote = WordXml.ParseInt32(foot.ChildVal("numStart")) ?? WordXml.ParseInt32(documentFoot.ChildVal("numStart")) ?? 1;
            if (!hasEndnotes || endRestart == "eachSect")
                nextEndnote = WordXml.ParseInt32(end.ChildVal("numStart")) ?? WordXml.ParseInt32(documentEnd.ChildVal("numStart")) ?? 1;
            footnotePageRestart = footRestart == "eachPage";
            endnotePageRestart = endRestart == "eachPage";
        }

        internal void ReportSeparatorMarkOmission() {
            if (separatorMarkOmissionReported || options == null) return;
            separatorMarkOmissionReported = true;
            AddNativeExportWarning(options, "native-footnote-separator-mark-omitted", "footnote separator",
                "The separator paragraph and its spacing are retained, but its rule mark is not painted because the selected font does not provide usable strike-band metrics.");
        }

        internal void ReportSeparatorDecorationOmission(string reason) {
            if (separatorDecorationOmissionReported || options == null) return;
            separatorDecorationOmissionReported = true;
            AddNativeExportWarning(options, "native-footnote-separator-decoration-omitted", "footnote separator",
                "The separator rule is retained, but an additional authored decoration is not painted. " + reason);
        }

        internal void ReportPlacementFallback(string? reason = null) {
            if (placementFallbackReported || options == null) return;
            placementFallbackReported = true;
            AddNativeExportWarning(options, "native-footnote-placement-approximated", "footnotes",
                "This footnote profile is placed in document flow after its section. Supported paragraph formatting is retained; page placement is approximated, and unsupported content is reported separately." +
                (reason == null ? string.Empty : " " + reason));
        }

        internal bool ContainsKey(NativeNoteKey key) => tokensById.ContainsKey(key);
        internal bool TryGetValue(NativeNoteKey key, out int token) => tokensById.TryGetValue(key, out token);
        internal string GetLabel(int token) => labelsByToken[token];

        internal string Add(NativeNoteKey key, bool endnote) {
            int number = (endnote ? nextEndnote : nextFootnote) ?? 1;
            if (endnote) { nextEndnote = checked(number + 1); hasEndnotes = true; }
            else { nextFootnote = checked(number + 1); hasFootnotes = true; }
            string format = endnote ? endnoteFormat : footnoteFormat;
            OfficeNumberStyle style;
            if (format == "lowerRoman") style = OfficeNumberStyle.LowerRoman;
            else if (format == "upperRoman") style = OfficeNumberStyle.UpperRoman;
            else if (format == "lowerLetter") style = OfficeNumberStyle.RepeatedLowerLetter;
            else if (format == "upperLetter") style = OfficeNumberStyle.RepeatedUpperLetter;
            else {
                style = OfficeNumberStyle.Decimal;
                if (format != "decimal" && options != null)
                    AddNativeExportWarning(options, "native-note-number-format-unsupported", endnote ? "endnote" : "footnote",
                        "The note numbering format is outside the supported decimal, Roman and letter formats; a decimal label was used.");
            }
            if ((endnote ? endnotePageRestart : footnotePageRestart) && !pageRestartReported && options != null) {
                AddNativeExportWarning(options, "native-note-page-restart-approximated", "notes",
                    "Page-based note numbering restart is not applied; note labels continue through the section.");
                pageRestartReported = true;
            }
            // Word permits a zero DOCX note start. Decimal displays 0; Roman and
            // letter formats leave that marker blank before continuing at 1.
            // Keep the shared formatter's positive page-number contract intact.
            string label = number == 0 ? (style == OfficeNumberStyle.Decimal ? "0" : string.Empty) :
                OfficeNumberFormatter.Format(number, style);
            if (label.Length > MaximumGeneratedLabelCharacters - generatedLabelCharacters)
                throw new InvalidDataException("The document's note labels exceed the cumulative character limit.");
            int token = tokensById.Count + 1;
            tokensById.Add(key, token);
            labelsByToken.Add(token, label);
            generatedLabelCharacters += label.Length;
            return label;
        }
    }
}
