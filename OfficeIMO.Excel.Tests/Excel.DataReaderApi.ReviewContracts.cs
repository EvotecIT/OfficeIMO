using System.Data.Common;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void OpenDataReader_RejectsMalformedWorksheetXmlAfterSheetData() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MalformedWorksheetTail.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "Ready");
                document.Save();
            }

            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            int worksheetEnd = worksheetXml.LastIndexOf("</", StringComparison.Ordinal);
            Assert.True(worksheetEnd >= 0);
            worksheetXml = worksheetXml.Insert(worksheetEnd, "<broken>");
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(worksheetXml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsNestedMarkupInSharedStringValue() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.NestedSharedStringValue.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "Ready");
                document.Save();
            }

            const string entryName = "xl/worksheets/sheet1.xml";
            XDocument worksheet = XDocument.Parse(
                Encoding.UTF8.GetString(ReadZipEntry(path, entryName)));
            XNamespace spreadsheet = worksheet.Root!.Name.Namespace;
            XElement cell = Assert.Single(
                worksheet.Descendants(spreadsheet + "c"),
                element => string.Equals(
                    (string?)element.Attribute("r"),
                    "A2",
                    StringComparison.Ordinal));
            XElement value = Assert.IsType<XElement>(cell.Element(spreadsheet + "v"));
            value.Add(new XElement(spreadsheet + "ext"));
            ReplaceZipEntry(
                path,
                entryName,
                Encoding.UTF8.GetBytes(worksheet.ToString(SaveOptions.DisableFormatting)));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));
            Assert.Contains("only text", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void OpenDataReader_RejectsMissingOpenXmlCellStylesBeforeDelivery(
        int invalidSheetIndex) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidOpenXmlCellStyle.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                for (int index = 0; index < 2; index++) {
                    ExcelSheet sheet = document.AddWorksheet($"Sheet{index + 1}");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(2, 1, $"Sheet {index + 1}");
                }
                document.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbookPart = package.WorkbookPart!;
                Sheet sheet = workbookPart.Workbook.Sheets!
                    .Elements<Sheet>()
                    .ElementAt(invalidSheetIndex);
                WorksheetPart worksheetPart = Assert.IsType<WorksheetPart>(
                    workbookPart.GetPartById(sheet.Id!.Value!));
                Cell dataCell = worksheetPart.Worksheet
                    .Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                dataCell.StyleIndex = uint.MaxValue;
                worksheetPart.Worksheet.Save();
            }

            if (invalidSheetIndex == 0) {
                InvalidDataException openException = Assert.Throws<InvalidDataException>(
                    () => ExcelDocument.OpenDataReader(path));
                Assert.Contains(
                    "missing cell style",
                    openException.Message,
                    StringComparison.OrdinalIgnoreCase);
                return;
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            InvalidDataException nextResultException = Assert.Throws<InvalidDataException>(
                () => reader.NextResult());
            Assert.Contains(
                "missing cell style",
                nextResultException.Message,
                StringComparison.OrdinalIgnoreCase);
            Assert.True(reader.IsClosed);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void CreateDataReader_RejectsMissingOpenXmlCellStylesBeforeDelivery() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidWrappedOpenXmlCellStyle.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "Ready");
                document.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                Cell dataCell = worksheetPart.Worksheet
                    .Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                dataCell.StyleIndex = uint.MaxValue;
                worksheetPart.Worksheet.Save();
            }

            using ExcelDocument loadedDocument = ExcelDocument.Load(path);
            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => loadedDocument.CreateDataReader());
            Assert.Contains(
                "missing cell style",
                exception.Message,
                StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void OpenDataReader_RejectsMissingOpenXmlSharedStringsBeforeDelivery(
        int invalidSheetIndex) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidOpenXmlSharedString.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                for (int index = 0; index < 2; index++) {
                    ExcelSheet sheet = document.AddWorksheet($"Sheet{index + 1}");
                    sheet.CellValue(1, 1, "Value");
                    sheet.CellValue(2, 1, $"Sheet {index + 1}");
                }
                document.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbookPart = package.WorkbookPart!;
                Sheet sheet = workbookPart.Workbook.Sheets!
                    .Elements<Sheet>()
                    .ElementAt(invalidSheetIndex);
                WorksheetPart worksheetPart = Assert.IsType<WorksheetPart>(
                    workbookPart.GetPartById(sheet.Id!.Value!));
                Cell dataCell = worksheetPart.Worksheet
                    .Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                dataCell.DataType = CellValues.SharedString;
                dataCell.CellValue = new CellValue(uint.MaxValue.ToString());
                worksheetPart.Worksheet.Save();
            }

            if (invalidSheetIndex == 0) {
                InvalidDataException openException = Assert.Throws<InvalidDataException>(
                    () => ExcelDocument.OpenDataReader(path));
                Assert.Contains(
                    "missing shared string",
                    openException.Message,
                    StringComparison.OrdinalIgnoreCase);
                return;
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            InvalidDataException nextResultException = Assert.Throws<InvalidDataException>(
                () => reader.NextResult());
            Assert.Contains(
                "missing shared string",
                nextResultException.Message,
                StringComparison.OrdinalIgnoreCase);
            Assert.True(reader.IsClosed);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbNextResultRejectsMissingSharedStringDuringDiscovery() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidSkippedSheetSharedString.{Guid.NewGuid():N}.xlsb");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Ready");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, "Never delivered");
                File.WriteAllBytes(path, document.ToBytes(ExcelFileFormat.Xlsb));
            }
            ReplaceXlsbTextCellWithSharedStringIndex(
                path,
                "xl/worksheets/sheet2.bin",
                occurrence: 2,
                uint.MaxValue);

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { InferSchema = false });
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => reader.NextResult());

            Assert.Contains("missing shared string", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(reader.IsClosed);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false, "missing relationship")]
    [InlineData(true, "external relationship")]
    public void OpenDataReader_RejectsInvalidOpenXmlWorksheetRelationship(
        bool external,
        string expectedMessage) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidWorksheetRelationship.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("First").CellValue(1, 1, "Ready");
                document.AddWorksheet("Second").CellValue(1, 1, "Never delivered");
                document.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbookPart = package.WorkbookPart!;
                Sheet second = workbookPart.Workbook.Sheets!.Elements<Sheet>().ElementAt(1);
                const string relationshipId = "rIdInvalidWorksheet";
                if (external) {
                    workbookPart.AddExternalRelationship(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        new Uri("https://example.invalid/sheet2.xml"),
                        relationshipId);
                }
                second.Id = relationshipId;
                workbookPart.Workbook.Save();
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains(expectedMessage, exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_RejectsSheetRelationshipToUnrelatedInternalPart() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.UnrelatedWorksheetPart.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("First").CellValue(1, 1, "Ready");
                document.AddWorksheet("Second").CellValue(1, 1, "Never delivered");
                document.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbookPart = package.WorkbookPart!;
                Sheet second = workbookPart.Workbook.Sheets!.Elements<Sheet>().ElementAt(1);
                WorkbookStylesPart stylesPart = Assert.IsType<WorkbookStylesPart>(
                    workbookPart.WorkbookStylesPart);
                second.Id = workbookPart.GetIdOfPart(stylesPart);
                workbookPart.Workbook.Save();
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("unsupported internal part", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_SkipsSupportedNonWorksheetSheetParts() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.NonWorksheetParts.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet data = document.AddWorksheet("Data");
                data.CellValue(1, 1, "Value");
                data.CellValue(2, 1, "Ready");
                document.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbookPart = package.WorkbookPart!;
                Sheets sheets = Assert.IsType<Sheets>(workbookPart.Workbook.Sheets);
                uint nextSheetId = sheets.Elements<Sheet>().Max(sheet => sheet.SheetId!.Value) + 1U;

                ChartsheetPart chartPart = workbookPart.AddNewPart<ChartsheetPart>();
                chartPart.Chartsheet = new Chartsheet();
                sheets.Append(new Sheet {
                    Name = "Chart",
                    SheetId = nextSheetId,
                    Id = workbookPart.GetIdOfPart(chartPart)
                });

                DialogsheetPart dialogPart = workbookPart.AddNewPart<DialogsheetPart>();
                dialogPart.DialogSheet = new DialogSheet();
                sheets.Append(new Sheet {
                    Name = "Dialog",
                    SheetId = nextSheetId + 1U,
                    Id = workbookPart.GetIdOfPart(dialogPart)
                });
                workbookPart.Workbook.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(1, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbClassifiesCustomDateStylesAfterAllStyleCollections() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.LateNumberFormats.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("styles-dates-formulas.xlsb"), path);
        try {
            MoveXlsbCollectionAfter(
                path,
                "xl/styles.bin",
                beginRecordType: 615,
                endRecordType: 616,
                targetEndRecordType: 618);

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { TreatDatesUsingNumberFormat = true });

            Assert.True(reader.Read());
            Assert.Equal(new DateTime(2024, 2, 29), reader.GetDateTime(0));
            Assert.Equal(new DateTime(2024, 2, 29), Assert.IsType<DateTime>(reader.GetValue(0)));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Reader_NumericGettersDoNotMaterializeOutOfRangeDateStyles() {
        using var memory = new MemoryStream();
        using (var document = ExcelDocument.Create(
            memory,
            new ExcelCreateOptions {
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            })) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            string[] headers = {
                "AsByte",
                "AsDecimal",
                "AsDouble",
                "AsFloat",
                "AsInt16",
                "AsInt32",
                "AsInt64"
            };
            for (int column = 1; column <= headers.Length; column++) {
                sheet.CellValue(1, column, headers[column - 1]);
                sheet.CellValue(2, column, column == 3 ? 1E100 : 42D);
                sheet.CellAt(2, column).SetNumberFormat("yyyy-mm-dd");
            }
        }

        using ExcelDocumentReader owner = ExcelDocumentReader.Open(
            memory.ToArray(),
            new ExcelReadOptions {
                TreatDatesUsingNumberFormat = true,
                NumericAsDecimal = true
            });
        using var reader = owner
            .GetSheet("Data")
            .ReadUsedRangeAsDataReader(schemaSampleRows: 0);

        Assert.True(reader.Read());
        Assert.Equal((byte)42, reader.GetByte(0));
        Assert.Equal(42M, reader.GetDecimal(1));
        Assert.Equal(1E100, reader.GetDouble(2));
        Assert.Equal(42F, reader.GetFloat(3));
        Assert.Equal((short)42, reader.GetInt16(4));
        Assert.Equal(42, reader.GetInt32(5));
        Assert.Equal(42L, reader.GetInt64(6));
    }

    [Fact]
    public void CreateDataReader_NumericGettersPreserveWrappedDocumentDateSerials() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Valid");
        sheet.CellValue(1, 2, "OutOfRange");
        sheet.CellValue(2, 1, 45351D);
        sheet.CellValue(2, 2, 1E100);
        sheet.CellAt(2, 1).SetNumberFormat("yyyy-mm-dd");
        sheet.CellAt(2, 2).SetNumberFormat("yyyy-mm-dd");

        using DbDataReader reader = document.CreateDataReader(
            new ExcelReadOptions { TreatDatesUsingNumberFormat = true });

        Assert.True(reader.Read());
        Assert.Equal(45351D, reader.GetDouble(0));
        Assert.Equal(new DateTime(2024, 2, 29), Assert.IsType<DateTime>(reader.GetValue(0)));
        Assert.Equal(1E100, reader.GetDouble(1));
    }

    [Theory]
    [InlineData("sharedStrings", "shared-string")]
    [InlineData("styles", "styles")]
    public void OpenDataReader_XlsbRejectsDuplicateSingletonWorkbookRelationships(
        string relationshipSuffix,
        string expectedMessage) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DuplicateWorkbookRelationship.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("styles-dates-formulas.xlsb"), path);
        try {
            DuplicateXlsbWorkbookRelationship(path, relationshipSuffix);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("multiple internal", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(expectedMessage, exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    private static void ReplaceXlsbTextCellWithSharedStringIndex(
        string path,
        string entryName,
        int occurrence,
        uint sharedStringIndex) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(input);
        using var output = new MemoryStream();
        int currentOccurrence = 0;
        bool replaced = false;
        foreach (XlsbRecord record in records) {
            byte[] data = record.Data;
            int recordType = record.Type;
            if ((record.Type == 6 || record.Type == 7)
                && ++currentOccurrence == occurrence) {
                Assert.True(data.Length >= 8);
                var replacement = new byte[12];
                Buffer.BlockCopy(data, 0, replacement, 0, 8);
                data = replacement;
                WriteUInt32LittleEndian(data, 8, sharedStringIndex);
                recordType = 7;
                replaced = true;
            }
            XlsbRecordWriter.Write(output, recordType, data);
        }

        Assert.True(replaced);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void MoveXlsbCollectionAfter(
        string path,
        string entryName,
        int beginRecordType,
        int endRecordType,
        int targetEndRecordType) {
        byte[] bytes = ReadZipEntry(path, entryName);
        using var input = new MemoryStream(bytes, writable: false);
        List<XlsbRecord> records = XlsbRecordReader.ReadAll(input).ToList();
        int beginIndex = records.FindIndex(record => record.Type == beginRecordType);
        int endIndex = records.FindIndex(
            beginIndex,
            record => record.Type == endRecordType);
        int targetEndIndex = records.FindIndex(record => record.Type == targetEndRecordType);
        Assert.True(beginIndex >= 0);
        Assert.True(endIndex >= beginIndex);
        Assert.True(targetEndIndex > endIndex);
        List<XlsbRecord> moved = records.GetRange(beginIndex, endIndex - beginIndex + 1);

        using var output = new MemoryStream();
        for (int index = 0; index < records.Count; index++) {
            if (index < beginIndex || index > endIndex) {
                XlsbRecord record = records[index];
                XlsbRecordWriter.Write(output, record.Type, record.Data);
            }
            if (index == targetEndIndex) {
                foreach (XlsbRecord record in moved) {
                    XlsbRecordWriter.Write(output, record.Type, record.Data);
                }
            }
        }

        ReplaceZipEntry(path, entryName, output.ToArray());
    }

    private static void DuplicateXlsbWorkbookRelationship(
        string path,
        string relationshipSuffix) {
        const string entryName = "xl/_rels/workbook.bin.rels";
        XDocument relationships = XDocument.Parse(
            Encoding.UTF8.GetString(ReadZipEntry(path, entryName)));
        XNamespace packageRelationships =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement source = Assert.Single(
            relationships.Descendants(packageRelationships + "Relationship"),
            element =>
                ((string?)element.Attribute("Type"))?.EndsWith(
                    "/" + relationshipSuffix,
                    StringComparison.Ordinal) == true);
        var duplicate = new XElement(source);
        duplicate.SetAttributeValue("Id", "rIdDuplicateSingleton");
        source.AddAfterSelf(duplicate);

        ReplaceZipEntry(
            path,
            entryName,
            Encoding.UTF8.GetBytes(relationships.ToString(SaveOptions.DisableFormatting)));
    }
}
