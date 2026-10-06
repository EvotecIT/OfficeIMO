using System.Collections.Generic;
using OfficeIMO.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Keeps reference identity distinct from labels that can repeat between note kinds or sections.</summary>
    private sealed class NativeNoteNumbering {
        private readonly Dictionary<long, int> tokensById = new();
        private readonly Dictionary<int, string> labelsByToken = new();
        private readonly W.Settings? settings;
        private readonly WordToPdfOptions? options;
        private int? nextFootnote;
        private int? nextEndnote;
        private bool hasFootnotes;
        private bool hasEndnotes;
        private W.NumberFormatValues footnoteFormat;
        private W.NumberFormatValues endnoteFormat;
        private bool pageRestart;
        private bool pageRestartReported;

        internal NativeNoteNumbering(WordDocument document, WordToPdfOptions? options) {
            settings = document._wordprocessingDocument?.MainDocumentPart?.DocumentSettingsPart?.Settings;
            this.options = options;
        }

        internal void BeginSection(WordSection section) {
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
            pageRestart = footRestart == W.RestartNumberValues.EachPage || endRestart == W.RestartNumberValues.EachPage;
        }

        internal bool ContainsKey(long key) => tokensById.ContainsKey(key);
        internal bool TryGetValue(long key, out int token) => tokensById.TryGetValue(key, out token);
        internal string GetLabel(int token) => labelsByToken[token];

        internal string Add(long key, bool endnote) {
            int number = (endnote ? nextEndnote : nextFootnote) ?? 1;
            if (endnote) { nextEndnote = checked(number + 1); hasEndnotes = true; }
            else { nextFootnote = checked(number + 1); hasFootnotes = true; }
            W.NumberFormatValues format = endnote ? endnoteFormat : footnoteFormat;
            OfficeNumberStyle style;
            if (format == W.NumberFormatValues.LowerRoman) style = OfficeNumberStyle.LowerRoman;
            else if (format == W.NumberFormatValues.UpperRoman) style = OfficeNumberStyle.UpperRoman;
            else if (format == W.NumberFormatValues.LowerLetter) style = OfficeNumberStyle.LowerLetter;
            else if (format == W.NumberFormatValues.UpperLetter) style = OfficeNumberStyle.UpperLetter;
            else {
                style = OfficeNumberStyle.Decimal;
                if (format != W.NumberFormatValues.Decimal && options != null)
                    AddNativeExportWarning(options, "NativeNoteNumberFormatUnsupported", endnote ? "endnote" : "footnote",
                        "The note numbering format is outside the supported decimal, Roman and letter formats; a decimal label was used.");
            }
            if (pageRestart && !pageRestartReported && options != null) {
                AddNativeExportWarning(options, "NativeNotePageRestartApproximated", "notes",
                    "Page-based note numbering restart requires note pagination; numbering continues through the section.");
                pageRestartReported = true;
            }
            string label = OfficeNumberFormatter.Format(number, style);
            int token = tokensById.Count + 1;
            tokensById.Add(key, token);
            labelsByToken.Add(token, label);
            return label;
        }
    }
}
