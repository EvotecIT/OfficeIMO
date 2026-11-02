namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkProjectionBudget {
    private readonly IWorkReadOptions _options;
    private int _tableCount;
    private int _imageCount;
    private int _drawableReferenceCount;
    private int _tableCatalogEntryCount;
    private int _tableDimensionEntryCount;
    private int _tableBorderEntryCount;
    private int _textItemCount;
    private int _textBoundaryCount;
    private long _textCharacterCount;
    private long _decodedImageByteCount;
    private long _projectedImageByteCount;
    private long _formulaRenderingOperations;

    internal IWorkProjectionBudget(IWorkReadOptions options) {
        _options = options;
    }

    internal int MaxProtobufFieldCount => _options.MaxProtobufFieldCount;

    internal int RemainingTableCatalogEntries =>
        _options.MaxTableCatalogEntries - _tableCatalogEntryCount;

    internal void AddTableCatalogEntries(int count) {
        if (count < 0 || _tableCatalogEntryCount > _options.MaxTableCatalogEntries - count) {
            throw new InvalidDataException(
                $"iWork table catalog entries exceed the configured source-wide limit of {_options.MaxTableCatalogEntries}.");
        }
        _tableCatalogEntryCount += count;
    }

    internal int RemainingTableDimensionEntries =>
        _options.MaxTableDimensionEntries - _tableDimensionEntryCount;

    internal void AddTableDimensionEntries(int count) {
        if (count < 0 || _tableDimensionEntryCount > _options.MaxTableDimensionEntries - count)
            throw IWorkProtobuf.LimitException(
                $"iWork table dimension entries exceed the configured source-wide limit of {_options.MaxTableDimensionEntries}.");
        _tableDimensionEntryCount += count;
    }

    internal void AddTable() {
        if (_tableCount >= _options.MaxProjectedTables) {
            throw new InvalidDataException(
                $"iWork table count exceeds the configured projection limit of {_options.MaxProjectedTables}.");
        }
        _tableCount++;
    }

    internal void AddTableBorderEntries(int count) {
        if (count < 0 || _tableBorderEntryCount > _options.MaxTableBorderEntries - count)
            throw IWorkProtobuf.LimitException(
                $"iWork table stroke declarations exceed the configured source-wide limit of {_options.MaxTableBorderEntries}.");
        _tableBorderEntryCount += count;
    }

    internal void AddImage() {
        if (_imageCount >= _options.MaxProjectedImages) {
            throw new InvalidDataException(
                $"iWork image count exceeds the configured projection limit of {_options.MaxProjectedImages}.");
        }
        _imageCount++;
    }

    internal void AddDrawableReferences(int count) {
        long maximum = Math.Min(int.MaxValue, (long)_options.MaxProjectedTables
            + _options.MaxProjectedImages + _options.MaxProjectedTextItems);
        if (count < 0 || _drawableReferenceCount > maximum - count) {
            throw IWorkProtobuf.LimitException(
                $"iWork drawable references exceed the configured combined projection limit of {maximum}.");
        }
        _drawableReferenceCount += count;
    }

    internal long RemainingDecodedImageBytes =>
        _options.MaxDecodedImageBytes - _decodedImageByteCount;

    internal void AddDecodedImageBytes(long count) {
        if (count < 0 || _decodedImageByteCount > _options.MaxDecodedImageBytes - count) {
            throw new InvalidDataException(
                $"Decoded image data exceeds the configured image limit of {_options.MaxDecodedImageBytes} bytes.");
        }
        _decodedImageByteCount += count;
    }

    internal void AddProjectedImageBytes(long count) {
        if (count < 0 || _projectedImageByteCount > _options.MaxProjectedImageBytes - count) {
            throw new InvalidDataException(
                $"Projected destination image data exceeds the configured limit of {_options.MaxProjectedImageBytes} bytes.");
        }
        _projectedImageByteCount += count;
    }

    internal void AddTextItem() {
        AddTextItems(1);
    }

    internal void AddTextItems(int count) {
        if (count < 0 || _textItemCount > _options.MaxProjectedTextItems - count) {
            throw new InvalidDataException(
                $"Text item count exceeds the configured projection limit of {_options.MaxProjectedTextItems}.");
        }
        _textItemCount += count;
    }

    internal void AddTextBoundaries(int count) {
        if (count < 0 || _textBoundaryCount > _options.MaxProjectedTextItems - count) {
            throw new InvalidDataException(
                $"Text attribute count exceeds the configured projection limit of {_options.MaxProjectedTextItems}.");
        }
        _textBoundaryCount += count;
    }

    internal void AddTextCharacters(int count) {
        if (count < 0 || _textCharacterCount > _options.MaxProjectedTextCharacters - count) {
            throw new InvalidDataException(
                $"Text character count exceeds the configured projection limit of {_options.MaxProjectedTextCharacters}.");
        }
        _textCharacterCount += count;
    }

    internal void AddFormulaRenderingOperations(long count) {
        if (count < 0
            || _formulaRenderingOperations > _options.MaxFormulaRenderingOperations - count) {
            throw new InvalidDataException(
                $"Formula rendering work exceeds the configured source-wide limit of {_options.MaxFormulaRenderingOperations} operations.");
        }
        _formulaRenderingOperations += count;
    }

    internal void AddTextContentUse(IWorkTextContent content, bool includeCharacters = false) {
        long count = content.Paragraphs.Count;
        long characterCount = 0;
        foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
            count += paragraph.Style.TabStops?.Count ?? 0;
            foreach (IWorkTextRun run in paragraph.Runs) {
                count++;
                if (includeCharacters) {
                    characterCount += run.Text.Length;
                    characterCount += run.Style.FontName?.Length ?? 0;
                    characterCount += run.Hyperlink?.Length ?? 0;
                }
                foreach (char character in run.Text) {
                    if (character == '\n') count++;
                }
            }
            if (includeCharacters) {
                characterCount += paragraph.ListLabel?.Length ?? 0;
                characterCount += paragraph.ListFontName?.Length ?? 0;
            }
        }
        if (count > int.MaxValue) {
            throw new InvalidDataException(
                $"Text item count exceeds the configured projection limit of {_options.MaxProjectedTextItems}.");
        }
        AddTextItems((int)count);
        if (includeCharacters) {
            if (characterCount > int.MaxValue) {
                throw new InvalidDataException(
                    $"Text character count exceeds the configured projection limit of {_options.MaxProjectedTextCharacters}.");
            }
            AddTextCharacters((int)characterCount);
        }
    }

    internal int MaxTextStyleInheritanceDepth => _options.MaxTextStyleInheritanceDepth;
}
