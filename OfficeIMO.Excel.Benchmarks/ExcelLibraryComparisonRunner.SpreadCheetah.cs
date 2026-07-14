using System.Data;
using System.Globalization;
using SpreadCheetah;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class ExcelLibraryComparisonRunner {
    private static int SpreadCheetahWriteDataReaderPlain(DataTable table)
        => ByteCount(SpreadCheetahWriteDataReaderPlainBytes(table));

    private static byte[] SpreadCheetahWriteDataReaderPlainBytes(DataTable table)
        => SpreadCheetahWriteDataReaderPlainBytesAsync(table).GetAwaiter().GetResult();

    private static async Task<byte[]> SpreadCheetahWriteDataReaderPlainBytesAsync(DataTable table) {
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream).ConfigureAwait(false);
        await spreadsheet.StartWorksheetAsync("Data").ConfigureAwait(false);

        using var reader = table.CreateDataReader();
        var cells = new DataCell[reader.FieldCount];
        for (int column = 0; column < cells.Length; column++) {
            cells[column] = new DataCell(reader.GetName(column));
        }

        await spreadsheet.AddRowAsync(cells).ConfigureAwait(false);
        while (reader.Read()) {
            for (int column = 0; column < cells.Length; column++) {
                cells[column] = CreateSpreadCheetahCell(reader.GetValue(column));
            }

            await spreadsheet.AddRowAsync(cells).ConfigureAwait(false);
        }

        await spreadsheet.FinishAsync().ConfigureAwait(false);
        return stream.ToArray();
    }

    private static DataCell CreateSpreadCheetahCell(object? value)
        => value switch {
            null or DBNull => default,
            string text => new DataCell(text),
            bool boolean => new DataCell(boolean),
            byte number => new DataCell(number),
            sbyte number => new DataCell(number),
            short number => new DataCell(number),
            ushort number => new DataCell(number),
            int number => new DataCell(number),
            uint number => new DataCell(number),
            long number => new DataCell(number),
            ulong number => new DataCell((decimal)number),
            float number => new DataCell(number),
            double number => new DataCell(number),
            decimal number => new DataCell(number),
            DateTime dateTime => new DataCell(dateTime),
            _ => new DataCell(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
        };
}
