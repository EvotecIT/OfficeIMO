using OfficeIMO.CSV;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Writes the rows of a tree source as CSV or TSV through the CSV engine: its UTF-8 row path when the rows hold UTF-8
/// cells, otherwise its streaming data-reader writer.
/// </summary>
internal static class CsvTableOutput {
    public static OfficeConversionReport Write(StructuredDataReader rows, Stream output, string source, bool tsv) {
        string target = tsv ? "TSV" : "CSV";
        DataConversionContext context = rows.Context;
        if (rows.TableCount > 1) {
            var names = new List<string>();
            do {
                names.Add(rows.TableName);
            } while (rows.NextResult());

            throw new InvalidDataException("The data holds " + names.Count.ToString(CultureInfo.InvariantCulture)
                + " tables, but " + target + " holds one. Choose one with Select: " + string.Join(", ", names) + ".");
        }

        CsvSaveOptions save = context.Options.CsvSave?.Clone() ?? (tsv ? CsvProfiles.CreateSaveOptions(CsvProfile.Tsv) : new CsvSaveOptions { NewLine = "\n" });
        if (context.Options.Culture != null) save.Culture = context.Options.Culture;
        if (rows.IsPositional) save.IncludeHeader = false;
#if NET8_0_OR_GREATER
        if (rows.HasUtf8Cells) {
            WriteUtf8(rows, output, save);
        } else {
            CsvDocument.WriteDataReader(output, rows, save, leaveOpen: true, context.CancellationToken);
        }
#else
        CsvDocument.WriteDataReader(output, rows, save, leaveOpen: true, context.CancellationToken);
#endif
        context.ThrowIfLossNotAllowed(source, target);
        return context.CreateReport(source, target);
    }

#if NET8_0_OR_GREATER
    private static void WriteUtf8(StructuredDataReader rows, Stream output, CsvSaveOptions save) {
        int count = rows.FieldCount;
        if (count <= 0) throw new InvalidOperationException("Data reader must expose at least one field.");
        var columns = new string[count];
        for (int i = 0; i < count; i++) columns[i] = rows.GetName(i);
        using CsvRowWriter writer = CsvRowWriter.CreateStream(output, save, leaveOpen: true);
        writer.WriteTextRows(columns, Array.Empty<string?[]>());
        if (rows.CanWriteJsonRows) {
            rows.WriteJsonRows(writer);
            return;
        }

        while (rows.Read()) writer.WriteUtf8Row(count, rows, static (reader, ordinal) => reader.GetUtf8(ordinal));
    }
#endif
}
