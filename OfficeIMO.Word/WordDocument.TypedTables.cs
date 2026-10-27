using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Adds an editable table from a single pass over typed properties, fields, or string-keyed dictionaries.</summary>
    public WordTable AddTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(
        IEnumerable<T> rows, OfficeTableOptions? options = null) {
        bool headers = options?.IncludeHeaders ?? true;
        var data = ObjectDataHelpers.ProjectTypedTable(rows, headers);
        int rowCount = checked(data.Rows.Count + (headers ? 1 : 0));
        if (rowCount == 0) throw new ArgumentException("A headerless table requires at least one row.", nameof(rows));
        var table = AddTable(rowCount, data.Headers.Length, WordTableStyle.TableGrid);
        table.Title = options?.Name;
        int offset = headers ? 1 : 0;
        if (headers) {
            table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
            for (int column = 0; column < data.Headers.Length; column++) GetOrCreateParagraph(table, 0, column).Text = data.Headers[column];
        }
        for (int row = 0; row < data.Rows.Count; row++)
            for (int column = 0; column < data.Headers.Length; column++)
                GetOrCreateParagraph(table, row + offset, column).Text = System.Convert.ToString(data.Rows[row][column], CultureInfo.CurrentCulture) ?? string.Empty;
        return table;
    }

    /// <summary>Adds an editable table from an existing rectangular matrix.</summary>
    public WordTable AddTable(object?[,] values, OfficeTableOptions? options = null) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        int rows = values.GetLength(0), columns = values.GetLength(1);
        if (rows == 0 || columns == 0) throw new ArgumentException("A table matrix requires at least one row and column.", nameof(values));
        var table = AddTable(rows, columns, WordTableStyle.TableGrid);
        table.Title = options?.Name;
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = options?.IncludeHeaders ?? true;
        for (int row = 0; row < rows; row++)
            for (int column = 0; column < columns; column++)
                GetOrCreateParagraph(table, row, column).Text = System.Convert.ToString(values[row, column], CultureInfo.CurrentCulture) ?? string.Empty;
        return table;
    }

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
}
