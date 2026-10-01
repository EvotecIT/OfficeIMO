using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private HtmlListState CreateListState(RtfListKind kind, IElement token) {
            IElement? firstItem = token.Children.FirstOrDefault(child => child.LocalName == "li");
            int id = firstItem == null ? NextAvailableListId() : ReadPositiveInteger(firstItem, "data-officeimo-rtf-list-id") ?? NextAvailableListId();
            int? definitionId = ResolveListDefinitionId(id) ?? (firstItem == null ? null : ReadPositiveInteger(firstItem, "data-officeimo-rtf-list-definition-id"));
            int depth = firstItem == null ? _lists.Count : ReadNonNegativeInteger(firstItem, "data-officeimo-rtf-list-level") ?? _lists.Count;
            int start = ReadInteger(token, "start") ?? 1;
            if (!definitionId.HasValue || !_document.ListDefinitions.Any(item => item.Id == definitionId.Value)) {
                int targetId = definitionId ?? id;
                while (_document.ListDefinitions.Any(item => item.Id == targetId)) targetId++;
                RtfListDefinition definition = _document.AddListDefinition(targetId);
                RtfListLevel level = definition.AddLevel(kind);
                for (int index = 0; index < depth; index++) level = definition.AddLevel(kind);
                level.StartAt = start;
                level.NumberFormat = GetAttribute(token, "type") switch { "I" => 1, "i" => 2, "A" => 3, "a" => 4, _ => kind == RtfListKind.Bullet ? 23 : 0 };
                definitionId = targetId;
                if (!_document.ListOverrides.Any(item => item.Id == id)) _document.AddListOverride(id, targetId);
            }
            var state = new HtmlListState(id, kind, depth, definitionId.Value, start);
            return state;
        }

        private int NextAvailableListId() {
            while (_document.ListOverrides.Any(item => item.Id == _nextListId)) _nextListId++;
            return _nextListId++;
        }

        private static int? ReadInteger(IElement token, string attributeName) => int.TryParse(GetAttribute(token, attributeName), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : null;

        private void ApplyListAttributes(IElement token) {
            RtfParagraph paragraph = EnsureParagraph();
            HtmlListState? state = _lists.Count == 0 ? null : _lists.Peek();
            if (state != null && ReadInteger(token, "value") is int value && value != state.NextValue && ReadPositiveInteger(token, "data-officeimo-rtf-list-id") == null) {
                state.Id = NextAvailableListId();
                RtfListLevelOverride levelOverride = _document.AddListOverride(state.Id, state.DefinitionId).AddLevelOverride();
                levelOverride.LevelIndex = state.Level;
                levelOverride.OverrideStartAt = true;
                levelOverride.StartAt = value;
                state.NextValue = value;
            }
            if (state != null) state.NextValue++;

            paragraph.ListKind = ReadListKind(token) ?? state?.Kind ?? RtfListKind.Bullet;
            paragraph.ListId = ReadPositiveInteger(token, "data-officeimo-rtf-list-id") ?? state?.Id ?? 1;
            paragraph.ListDefinitionId = ReadPositiveInteger(token, "data-officeimo-rtf-list-definition-id") ?? ResolveListDefinitionId(paragraph.ListId);
            paragraph.ListLevel = ReadNonNegativeInteger(token, "data-officeimo-rtf-list-level") ?? state?.Level ?? 0;

            string? listText = GetAttribute(token, "data-officeimo-rtf-list-text");
            if (listText != null) {
                paragraph.SetListText(listText);
            }
        }

        private static int? ReadPositiveInteger(IElement token, string attributeName) {
            int? value = ReadNonNegativeInteger(token, attributeName);
            return value.HasValue && value.Value > 0 ? value.Value : null;
        }

        private static int? ReadNonNegativeInteger(IElement token, string attributeName) {
            string? value = GetAttribute(token, attributeName);
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
                ? parsed
                : null;
        }

        private static RtfListKind? ReadListKind(IElement token) {
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

        private int? ResolveListDefinitionId(int? listId) {
            if (!listId.HasValue) {
                return null;
            }

            return _document.ListOverrides.FirstOrDefault(item => item.Id == listId.Value)?.ListId;
        }

        private sealed class HtmlListState {
            internal HtmlListState(int id, RtfListKind kind, int level, int definitionId, long start) {
                Id = id;
                Kind = kind;
                Level = level;
                DefinitionId = definitionId;
                NextValue = start;
            }

            internal int Id { get; set; }

            internal int DefinitionId { get; }

            internal long NextValue { get; set; }

            internal RtfListKind Kind { get; }

            internal int Level { get; }
        }
    }
}
