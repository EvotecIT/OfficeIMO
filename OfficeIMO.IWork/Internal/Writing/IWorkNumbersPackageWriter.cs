using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.IWork.Internal;

internal static class IWorkNumbersPackageWriter {
    internal static byte[] Write(IWorkNumbersDocument document, IWorkNumbersWriteOptions limits, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (document.Sheets.Count == 0) throw new InvalidOperationException("A native Numbers spreadsheet requires at least one sheet.");
        if (document.Sheets.Count > limits.MaximumSheets) throw new InvalidDataException("The native Numbers sheet limit was exceeded.");
        var sheets = new List<(string Name, IWorkNativeTableData[] Tables)>();
        long tableCount = 0, cellCount = 0, textCount = 0;
        foreach (IWorkNumbersSheetBuilder sheet in document.Sheets) {
            token.ThrowIfCancellationRequested();
            if (sheet.Tables.Count == 0) throw new InvalidOperationException("Each native Numbers sheet requires at least one table.");
            tableCount += sheet.Tables.Count;
            if (tableCount > limits.MaximumTables) throw new InvalidDataException("The native Numbers table limit was exceeded.");
            var tables = sheet.Tables.Select(table => table.Data).ToArray();
            textCount += sheet.Name.Length;
            foreach (IWorkNativeTableData table in tables) {
                token.ThrowIfCancellationRequested();
                cellCount += (long)table.RowCount * table.ColumnCount;
                if (cellCount > limits.MaximumTableCells) throw new InvalidDataException("The native Numbers table-cell limit was exceeded.");
                textCount += table.Name.Length;
                foreach (IWorkTableCell cell in table.Cells) {
                    token.ThrowIfCancellationRequested();
                    if (cell.Kind == IWorkCellKind.Text) textCount += ((string)cell.Value!).Length;
                    if (textCount > limits.MaximumTextCharacters) throw new InvalidDataException("The native Numbers text-character limit was exceeded.");
                }
                if (textCount > limits.MaximumTextCharacters) throw new InvalidDataException("The native Numbers text-character limit was exceeded.");
            }
            sheets.Add((sheet.Name, tables));
        }
        IWorkProtoWriter P() => new(limits.MaximumPackageBytes, token);
        var hashInput = P();
        foreach (var sheet in sheets) {
            var row = P().String(1, sheet.Name);
            foreach (IWorkNativeTableData table in sheet.Tables) row.Message(2, table.ModelFields(limits.MaximumPackageBytes, token));
            hashInput.Message(1, row);
        }
        using var sha = SHA256.Create();
        var graph = new IWorkNumbersArchiveBuilder(limits.MaximumPackageBytes, token, sha.ComputeHash(hashInput.ToArray()));
        var archives = graph.Build(sheets);
        return IWorkNativePackageWriter.Write(archives.Document, archives.Metadata, graph.DocumentIdentifier,
            limits.MaximumPackageBytes, token);
    }
}
