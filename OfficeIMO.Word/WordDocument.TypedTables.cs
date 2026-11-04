using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Adds an editable table from a single pass over typed properties, fields, or string-keyed dictionaries.</summary>
    /// <remarks><see cref="OfficeTableOptions.Name"/> becomes the table's accessible title. <see cref="OfficeTableOptions.Caption"/> adds a visible paragraph immediately before the table. Pass <see cref="WordTableOptions"/> to choose the table style and layout mode.</remarks>
    public WordTable AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options = null) {
        bool headers = options?.IncludeHeaders ?? true;
        var data = ObjectDataHelpers.ProjectTypedTable(rows, headers, options: options);
        int rowCount = checked(data.Rows.Count + (headers ? 1 : 0));
        if (rowCount == 0) throw new ArgumentException("A headerless table requires at least one row.", nameof(rows));
        var table = CreateOptionsTable(rowCount, data.Headers.Length, options);
        int offset = headers ? 1 : 0;
        if (headers) {
            for (int column = 0; column < data.Headers.Length; column++) GetOrCreateParagraph(table, 0, column).Text = data.Headers[column];
        }
        for (int row = 0; row < data.Rows.Count; row++)
            for (int column = 0; column < data.Headers.Length; column++)
                GetOrCreateParagraph(table, row + offset, column).Text = System.Convert.ToString(data.Rows[row][column], CultureInfo.CurrentCulture) ?? string.Empty;
        ApplyTableAlignments(table, options);
        return table;
    }

    /// <summary>Adds an editable table from an existing rectangular matrix.</summary>
    public WordTable AddTable(object?[,] values, OfficeTableOptions? options = null) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        ObjectDataHelpers.ValidateTableAlignments(options);
        int rows = values.GetLength(0), columns = values.GetLength(1);
        if (rows == 0 || columns == 0) throw new ArgumentException("A table matrix requires at least one row and column.", nameof(values));
        var table = CreateOptionsTable(rows, columns, options);
        for (int row = 0; row < rows; row++)
            for (int column = 0; column < columns; column++)
                GetOrCreateParagraph(table, row, column).Text = System.Convert.ToString(values[row, column], CultureInfo.CurrentCulture) ?? string.Empty;
        ApplyTableAlignments(table, options);
        return table;
    }

    /// <summary>Adds existing array rows as table cells. The first row is the header when headers are enabled.</summary>
    public WordTable AddTable(IEnumerable<string[]> rows, OfficeTableOptions? options = null) =>
        AddTable((IEnumerable<IReadOnlyList<string>>)rows, options);

    /// <summary>Adds existing rows of text. The first row is the header when headers are enabled.</summary>
    public WordTable AddTable(IEnumerable<IReadOnlyList<string>> rows, OfficeTableOptions? options = null) {
        if (rows == null) throw new ArgumentNullException(nameof(rows));
        var data = rows.Select(row => row?.ToArray() ?? throw new ArgumentException("Table rows must not be null.", nameof(rows))).ToArray();
        if (data.Length == 0 || data[0].Length == 0 || data.Any(row => row.Length != data[0].Length))
            throw new ArgumentException("Text rows must form a non-empty rectangular table.", nameof(rows));
        var matrix = new object?[data.Length, data[0].Length];
        for (int row = 0; row < data.Length; row++)
            for (int column = 0; column < data[row].Length; column++) matrix[row, column] = data[row][column];
        return AddTable(matrix, options);
    }

    /// <summary>Creates the native table for an options-based overload: caption paragraph, style, layout, title and header-row flags.</summary>
    private WordTable CreateOptionsTable(int rows, int columns, OfficeTableOptions? options) {
        var wordOptions = options as WordTableOptions;
        if (options?.Caption != null) AddParagraph(options.Caption);
        var table = AddTable(rows, columns, wordOptions?.Style ?? WordTableStyle.OfficeIMOModern);
        if (wordOptions?.LayoutMode is { } layoutMode) table.LayoutMode = layoutMode;
        table.Title = options?.Name;
        bool headers = options?.IncludeHeaders ?? true;
        table.ConditionalFormattingFirstRow = headers;
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = headers;
        return table;
    }
}
