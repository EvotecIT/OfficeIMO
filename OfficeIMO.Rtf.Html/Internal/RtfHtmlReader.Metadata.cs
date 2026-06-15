using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private bool TryReadDocumentMetadata(HtmlToken token) {
            switch (token.Value) {
                case "head":
                    _headDepth++;
                    return true;
                case "title":
                    if (_headDepth > 0) {
                        _titleDepth++;
                        _titleText ??= new StringBuilder();
                        return true;
                    }

                    break;
                case "meta":
                    if (_headDepth > 0) {
                        ReadMeta(token);
                        return true;
                    }

                    break;
            }

            return _headDepth > 0;
        }

        private bool EndDocumentMetadata(string name) {
            switch (name) {
                case "title":
                    if (_titleDepth > 0) {
                        _titleDepth--;
                        if (_titleDepth == 0 && _titleText != null && !string.IsNullOrWhiteSpace(_titleText.ToString()) && string.IsNullOrWhiteSpace(_document.Info.Title)) {
                            _document.Info.Title = _titleText.ToString().Trim();
                        }

                        return true;
                    }

                    break;
                case "head":
                    if (_headDepth > 0) {
                        _headDepth--;
                        return true;
                    }

                    break;
            }

            return _headDepth > 0;
        }

        private bool AppendDocumentMetadataText(string text) {
            if (_titleDepth > 0) {
                _titleText ??= new StringBuilder();
                _titleText.Append(text);
                return true;
            }

            return _headDepth > 0;
        }

        private void ReadMeta(HtmlToken token) {
            string? name = GetAttribute(token, "name");
            string? content = GetAttribute(token, "content");
            if (string.IsNullOrWhiteSpace(name)) {
                return;
            }

            switch (name!.Trim().ToLowerInvariant()) {
                case "author":
                case "officeimo-rtf-author":
                    _document.Info.Author = EmptyToNull(content);
                    break;
                case "keywords":
                case "officeimo-rtf-keywords":
                    _document.Info.Keywords = EmptyToNull(content);
                    break;
                case "description":
                case "officeimo-rtf-subject":
                    _document.Info.Subject = EmptyToNull(content);
                    break;
                case "officeimo-rtf-title":
                    _document.Info.Title = EmptyToNull(content);
                    break;
                case "officeimo-rtf-manager":
                    _document.Info.Manager = EmptyToNull(content);
                    break;
                case "officeimo-rtf-company":
                    _document.Info.Company = EmptyToNull(content);
                    break;
                case "officeimo-rtf-operator":
                    _document.Info.Operator = EmptyToNull(content);
                    break;
                case "officeimo-rtf-category":
                    _document.Info.Category = EmptyToNull(content);
                    break;
                case "officeimo-rtf-comments":
                    _document.Info.Comments = EmptyToNull(content);
                    break;
                case "officeimo-rtf-hyperlink-base":
                    _document.Info.HyperlinkBase = EmptyToNull(content);
                    break;
                case "officeimo-rtf-created":
                    _document.Info.Created = ParseTimestamp(content);
                    break;
                case "officeimo-rtf-revised":
                    _document.Info.Revised = ParseTimestamp(content);
                    break;
                case "officeimo-rtf-printed":
                    _document.Info.Printed = ParseTimestamp(content);
                    break;
                case "officeimo-rtf-backed-up":
                    _document.Info.BackedUp = ParseTimestamp(content);
                    break;
                case "officeimo-rtf-editing-minutes":
                    _document.Info.EditingMinutes = ParseInteger(content);
                    break;
                case "officeimo-rtf-pages":
                    _document.Info.NumberOfPages = ParseInteger(content);
                    break;
                case "officeimo-rtf-words":
                    _document.Info.NumberOfWords = ParseInteger(content);
                    break;
                case "officeimo-rtf-characters":
                    _document.Info.NumberOfCharacters = ParseInteger(content);
                    break;
                case "officeimo-rtf-characters-with-spaces":
                    _document.Info.NumberOfCharactersWithSpaces = ParseInteger(content);
                    break;
                case "officeimo-rtf-internal-version":
                    _document.Info.InternalVersion = ParseInteger(content);
                    break;
                case "officeimo-rtf-header-footer":
                    ReadHeaderFooter(token, content);
                    break;
                case "officeimo-rtf-colors":
                    ApplyColorTable(RtfHtmlMetadataCodec.Decode(content));
                    break;
                case "officeimo-rtf-document-layout":
                    ApplyDocumentLayout(RtfHtmlMetadataCodec.Decode(content));
                    break;
                case "officeimo-rtf-document-settings":
                    ApplyDocumentSettings(RtfHtmlMetadataCodec.Decode(content));
                    break;
            }
        }

        private void ReadHeaderFooter(HtmlToken token, string? encodedContent) {
            if (!TryParseHeaderFooterKind(GetAttribute(token, "data-officeimo-rtf-kind"), out RtfHeaderFooterKind kind)) {
                return;
            }

            string? html = DecodeString(encodedContent);
            if (string.IsNullOrEmpty(html)) {
                return;
            }

            RtfHeaderFooter headerFooter = _document.AddHeaderFooter(kind);
            RtfDocument contentDocument = html!.ToRtfDocumentFromHtml();
            foreach (RtfParagraph paragraph in contentDocument.Paragraphs) {
                RtfParagraph target = headerFooter.AddParagraph();
                CopyParagraphInlines(paragraph, target, contentDocument);
            }
        }

        private static bool TryParseHeaderFooterKind(string? value, out RtfHeaderFooterKind kind) {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant()) {
                case "header":
                    kind = RtfHeaderFooterKind.Header;
                    return true;
                case "left-header":
                case "headerl":
                    kind = RtfHeaderFooterKind.LeftHeader;
                    return true;
                case "right-header":
                case "headerr":
                    kind = RtfHeaderFooterKind.RightHeader;
                    return true;
                case "first-header":
                case "headerf":
                    kind = RtfHeaderFooterKind.FirstHeader;
                    return true;
                case "footer":
                    kind = RtfHeaderFooterKind.Footer;
                    return true;
                case "left-footer":
                case "footerl":
                    kind = RtfHeaderFooterKind.LeftFooter;
                    return true;
                case "right-footer":
                case "footerr":
                    kind = RtfHeaderFooterKind.RightFooter;
                    return true;
                case "first-footer":
                case "footerf":
                    kind = RtfHeaderFooterKind.FirstFooter;
                    return true;
                default:
                    kind = RtfHeaderFooterKind.Header;
                    return false;
            }
        }

        private static string? EmptyToNull(string? value) {
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private static int? ParseInteger(string? value) {
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
                ? parsed
                : null;
        }

        private static DateTime? ParseTimestamp(string? value) {
            return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime parsed)
                ? parsed
                : null;
        }
    }
}
