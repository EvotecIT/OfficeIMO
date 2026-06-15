namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private bool TryReadGeneratedText(IElement token) {
            string? value = GetAttribute(token, "data-officeimo-rtf-generated-text");
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            EnsureInlineParagraph().AddGeneratedText(ReadGeneratedTextKind(value!));
            return true;
        }

        private static RtfGeneratedTextKind ReadGeneratedTextKind(string value) {
            switch (value.Trim().ToLowerInvariant()) {
                case "section-number":
                case "section":
                case "sectnum":
                    return RtfGeneratedTextKind.SectionNumber;
                case "current-date":
                case "date":
                case "chdate":
                    return RtfGeneratedTextKind.CurrentDate;
                case "current-date-long":
                case "date-long":
                case "chdpl":
                    return RtfGeneratedTextKind.CurrentDateLong;
                case "current-date-abbreviated":
                case "date-abbreviated":
                case "chdpa":
                    return RtfGeneratedTextKind.CurrentDateAbbreviated;
                case "current-time":
                case "time":
                case "chtime":
                    return RtfGeneratedTextKind.CurrentTime;
                default:
                    return RtfGeneratedTextKind.PageNumber;
            }
        }
    }
}
