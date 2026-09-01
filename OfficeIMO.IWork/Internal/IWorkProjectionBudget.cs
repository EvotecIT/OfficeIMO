namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkProjectionBudget {
    private readonly IWorkReadOptions _options;
    private int _tableCount;
    private int _imageCount;
    private int _drawableReferenceCount;
    private int _tableCatalogEntryCount;
    private int _textItemCount;
    private int _textBoundaryCount;
    private long _textCharacterCount;
    private long _decodedImageByteCount;
    private long _projectedImageByteCount;
    private long _formulaRenderingOperations;

    internal IWorkProjectionBudget(IWorkReadOptions options) {
        _options = options;
    }

    internal int MaximumProtobufFieldCount => _options.MaximumProtobufFieldCount;

    internal int RemainingTableCatalogEntries =>
        _options.MaximumTableCatalogEntries - _tableCatalogEntryCount;

    internal void AddTableCatalogEntries(int count) {
        if (count < 0 || _tableCatalogEntryCount > _options.MaximumTableCatalogEntries - count) {
            throw new InvalidDataException(
                $"iWork table catalog entries exceed the configured source-wide limit of {_options.MaximumTableCatalogEntries}.");
        }
        _tableCatalogEntryCount += count;
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

    internal void AddDrawableReferences(int count) {
        long maximum = Math.Min(int.MaxValue, (long)_options.MaximumProjectedTables
            + _options.MaximumProjectedImages + _options.MaximumProjectedTextItems);
        if (count < 0 || _drawableReferenceCount > maximum - count) {
            throw new InvalidDataException(
                $"iWork drawable references exceed the configured combined projection limit of {maximum}.");
        }
        _drawableReferenceCount += count;
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

    internal void AddProjectedImageBytes(long count) {
        if (count < 0 || _projectedImageByteCount > _options.MaximumProjectedImageBytes - count) {
            throw new InvalidDataException(
                $"Projected destination image data exceeds the configured limit of {_options.MaximumProjectedImageBytes} bytes.");
        }
        _projectedImageByteCount += count;
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

    internal void AddFormulaRenderingOperations(long count) {
        if (count < 0
            || _formulaRenderingOperations > _options.MaximumFormulaRenderingOperations - count) {
            throw new InvalidDataException(
                $"Formula rendering work exceeds the configured source-wide limit of {_options.MaximumFormulaRenderingOperations} operations.");
        }
        _formulaRenderingOperations += count;
    }

    internal void AddTextContentUse(IWorkTextContent content, bool includeCharacters = false) {
        long count = content.Paragraphs.Count;
        long characterCount = 0;
        foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
            foreach (IWorkTextRun run in paragraph.Runs) {
                count++;
                if (includeCharacters) characterCount += run.Text.Length;
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
        if (includeCharacters) {
            if (characterCount > int.MaxValue) {
                throw new InvalidDataException(
                    $"Text character count exceeds the configured projection limit of {_options.MaximumProjectedTextCharacters}.");
            }
            AddTextCharacters((int)characterCount);
        }
    }

    internal int MaximumTextStyleInheritanceDepth => _options.MaximumTextStyleInheritanceDepth;
}
