using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

public sealed partial class ImportOfficeCsvCommand {
    /// <summary>
    /// Parses text as the Text parameter sets do, writing through another command's runtime, so ConvertFrom-OfficeCsv
    /// shares this command's options, validation and output shapes.
    /// </summary>
    internal void ImportText(ICommandRuntime runtime, string text) {
        CommandRuntime = runtime;
        Text = text;
        BeginProcessing();
        ApplyCultureDelimiter();
        WriteText();
    }

    /// <summary>Stops a running <see cref="ImportText"/>.</summary>
    internal void Cancel() => _cancellation.Cancel();

    private void WriteText() {
        _rowWriter.Reset();
        _parseErrors.Clear();
        if (As == OfficeCsvOutput.Document || Mode != CsvReadMode.Stream || RequiresMaterializedRows()) {
            WriteDocumentRows(CsvDocument.Parse(Text, CreateLoadOptions()));
        } else {
            var reader = CsvDocument.OpenReader(new StringReader(Text), ConfigureReaderOptions(CreateReaderInputOptions()));
            if (_asDataReader) WriteObject(PSObject.AsPSObject(reader), enumerateCollection: false);
            else using (reader) {
                if (_asDataTable) {
                    var table = new System.Data.DataTable();
                    table.Load(reader);
                    WriteObject(PSObject.AsPSObject(table), enumerateCollection: false);
                } else _rowWriter.WriteDataReaderRows(reader, _asHashtable, this);
            }
        }
        WriteCollectedParseErrors("Text");
    }
}
