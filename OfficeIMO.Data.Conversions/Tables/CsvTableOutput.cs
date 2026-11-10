using OfficeIMO.CSV;

namespace OfficeIMO.Data.Conversions;

/// <summary>Writes the rows of a tree source as CSV through the CSV engine's streaming data-reader writer.</summary>
internal static class CsvTableOutput {
    public static OfficeConversionReport Write(StructuredDataReader rows, Stream output, string source) {
        DataConversionContext context = rows.Context;
        if (rows.TableCount > 1) {
            var names = new List<string>();
            do {
                names.Add(rows.TableName);
            } while (rows.NextResult());

            throw new InvalidDataException("The data holds " + names.Count.ToString(CultureInfo.InvariantCulture)
                + " tables, but CSV holds one. Choose one with Select: " + string.Join(", ", names) + ".");
        }

        CsvSaveOptions save = context.Options.CsvSave?.Clone() ?? new CsvSaveOptions { NewLine = "\n" };
        if (context.Options.Culture != null) save.Culture = context.Options.Culture;
        if (rows.IsPositional) save.IncludeHeader = false;
        CsvDocument.WriteDataReader(output, rows, save, leaveOpen: true, context.CancellationToken);
        context.ThrowIfLossNotAllowed(source, "CSV");
        return context.CreateReport(source, "CSV");
    }
}
