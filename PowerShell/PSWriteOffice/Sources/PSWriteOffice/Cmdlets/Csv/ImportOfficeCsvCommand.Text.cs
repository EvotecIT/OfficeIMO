using System.IO;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

public sealed partial class ImportOfficeCsvCommand {
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
