namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkProjectionBudget {
    private readonly IWorkReadOptions _options;
    private int _tableCount;
    private int _imageCount;
    private int _textItemCount;
    private int _textBoundaryCount;
    private long _textCharacterCount;
    private long _decodedImageByteCount;

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

    internal void AddImage() {
        if (_imageCount >= _options.MaximumProjectedImages) {
            throw new InvalidDataException(
                $"iWork image count exceeds the configured projection limit of {_options.MaximumProjectedImages}.");
        }
        _imageCount++;
    }

    internal long RemainingDecodedImageBytes =>
        _options.MaximumPackageBytes - _decodedImageByteCount;

    internal void AddDecodedImageBytes(long count) {
        if (count < 0 || _decodedImageByteCount > _options.MaximumPackageBytes - count) {
            throw new InvalidDataException(
                $"Decoded image data exceeds the configured package limit of {_options.MaximumPackageBytes} bytes.");
        }
        _decodedImageByteCount += count;
    }

    internal void AddTextItem() {
        AddTextItems(1);
    }

    internal void AddTextItems(int count) {
        if (count < 0 || _textItemCount > _options.MaximumProjectedTextItems - count) {
            throw new InvalidDataException(
                $"Text item count exceeds the configured projection limit of {_options.MaximumProjectedTextItems}.");
        }
        _textItemCount += count;
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

    internal void AddTextContentUse(IWorkTextContent content) {
        long count = content.Paragraphs.Count;
        foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
            foreach (IWorkTextRun run in paragraph.Runs) {
                count++;
                foreach (char character in run.Text) {
                    if (character == '\n') count++;
                }
            }
        }
        if (count > int.MaxValue) {
            throw new InvalidDataException(
                $"Text item count exceeds the configured projection limit of {_options.MaximumProjectedTextItems}.");
        }
        AddTextItems((int)count);
    }

    internal int MaximumTextStyleInheritanceDepth => _options.MaximumTextStyleInheritanceDepth;
}
