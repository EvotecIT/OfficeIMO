namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkProjectionBudget {
    private readonly IWorkReadOptions _options;
    private int _tableCount;
    private int _textItemCount;
    private int _textBoundaryCount;
    private long _textCharacterCount;

    internal IWorkProjectionBudget(IWorkReadOptions options) {
        _options = options;
    }

    internal void AddTable() {
        if (_tableCount >= _options.MaximumProjectedTables) {
            throw new InvalidDataException(
                $"iWork table count exceeds the configured projection limit of {_options.MaximumProjectedTables}.");
        }
        _tableCount++;
    }

    internal void AddTextItem() {
        if (_textItemCount >= _options.MaximumProjectedTextItems) {
            throw new InvalidDataException(
                $"Text item count exceeds the configured projection limit of {_options.MaximumProjectedTextItems}.");
        }
        _textItemCount++;
    }

    internal void AddTextBoundaries(int count) {
        if (count < 0 || _textBoundaryCount > _options.MaximumProjectedTextItems - count) {
            throw new InvalidDataException(
                $"Text attribute count exceeds the configured projection limit of {_options.MaximumProjectedTextItems}.");
        }
        _textBoundaryCount += count;
    }

    internal void AddTextCharacters(int count) {
        if (count < 0 || _textCharacterCount > _options.MaximumProjectedTextCharacters - count) {
            throw new InvalidDataException(
                $"Text character count exceeds the configured projection limit of {_options.MaximumProjectedTextCharacters}.");
        }
        _textCharacterCount += count;
    }

    internal int MaximumTextStyleInheritanceDepth => _options.MaximumTextStyleInheritanceDepth;
}
