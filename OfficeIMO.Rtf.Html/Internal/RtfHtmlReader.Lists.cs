using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private HtmlListState CreateListState(RtfListKind kind) {
            var state = new HtmlListState(_nextListId++, kind, _lists.Count);
            return state;
        }

        private void ApplyListAttributes(HtmlToken token) {
            RtfParagraph paragraph = EnsureParagraph();
            HtmlListState? state = _lists.Count == 0 ? null : _lists.Peek();

            paragraph.ListKind = ReadListKind(token) ?? state?.Kind ?? RtfListKind.Bullet;
            paragraph.ListId = ReadPositiveInteger(token, "data-officeimo-rtf-list-id") ?? state?.Id ?? 1;
            paragraph.ListDefinitionId = ReadPositiveInteger(token, "data-officeimo-rtf-list-definition-id");
            paragraph.ListLevel = ReadNonNegativeInteger(token, "data-officeimo-rtf-list-level") ?? state?.Level ?? 0;

            string? listText = GetAttribute(token, "data-officeimo-rtf-list-text");
            if (listText != null) {
                paragraph.SetListText(listText);
            }
        }

        private static int? ReadPositiveInteger(HtmlToken token, string attributeName) {
            int? value = ReadNonNegativeInteger(token, attributeName);
            return value.HasValue && value.Value > 0 ? value.Value : null;
        }

        private static int? ReadNonNegativeInteger(HtmlToken token, string attributeName) {
            string? value = GetAttribute(token, attributeName);
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
                ? parsed
                : null;
        }

        private static RtfListKind? ReadListKind(HtmlToken token) {
            string? value = GetAttribute(token, "data-officeimo-rtf-list-kind");
            switch (value?.Trim().ToLowerInvariant()) {
                case "bullet":
                case "ul":
                    return RtfListKind.Bullet;
                case "decimal":
                case "number":
                case "numbered":
                case "ol":
                    return RtfListKind.Decimal;
                default:
                    return null;
            }
        }

        private sealed class HtmlListState {
            internal HtmlListState(int id, RtfListKind kind, int level) {
                Id = id;
                Kind = kind;
                Level = level;
            }

            internal int Id { get; }

            internal RtfListKind Kind { get; }

            internal int Level { get; }
        }
    }
}
