namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkProjectionBudget {
    private readonly IWorkReadOptions _options;
    private int _tableCount;
    private int _textItemCount;

    internal IWorkProjectionBudget(IWorkReadOptions options) {
        _options = options;
    }

    internal void AddTable() {
        if (_tableCount >= _options.MaximumProjectedTables) {
            throw new InvalidDataException(
                $"Numbers table count exceeds the configured projection limit of {_options.MaximumProjectedTables}.");
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
}
