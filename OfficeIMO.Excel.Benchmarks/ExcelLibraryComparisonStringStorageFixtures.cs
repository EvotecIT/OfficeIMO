using System.Globalization;
using System.IO.Compression;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class ExcelLibraryComparisonRunner {
    private static byte[] CreateInlineStringWorkbookBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, autoSave: true)) {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            WriteInlineStringFixtureWorksheet(worksheetPart, rowCount);
            workbookPart.Workbook = new Workbook(
                new Sheets(
                    new Sheet {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = "Strings"
                    }));
            workbookPart.Workbook.Save();
        }

        byte[] workbookBytes = GetValidatedWorkbookFixtureBytes(stream, "inline-string read");
        ValidateStringStorageFixture(workbookBytes, rowCount, expectedCellType: "inlineStr", expectedUniqueSharedStrings: null);
        return workbookBytes;
    }

    private static byte[] CreateSharedStringWorkbookBytes(int rowCount) {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, autoSave: true)) {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            SharedStringTablePart sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();

            WriteSharedStringFixtureTable(sharedStringPart, rowCount);
            WriteSharedStringFixtureWorksheet(worksheetPart, rowCount);

            workbookPart.Workbook = new Workbook(
                new Sheets(
                    new Sheet {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = "Strings"
                    }));
            workbookPart.Workbook.Save();
        }

        byte[] workbookBytes = GetValidatedWorkbookFixtureBytes(stream, "shared-string-table read");
        ValidateStringStorageFixture(workbookBytes, rowCount, expectedCellType: "s", expectedUniqueSharedStrings: rowCount + 38);
        return workbookBytes;
    }

    private static void WriteSharedStringFixtureTable(SharedStringTablePart part, int rowCount) {
        int uniqueCount = checked(rowCount + 38);
        using OpenXmlWriter writer = OpenXmlWriter.Create(part);
        writer.WriteStartElement(new SharedStringTable {
            Count = checked((uint)(rowCount * 3)),
            UniqueCount = checked((uint)uniqueCount)
        });

        for (int index = 0; index < 12; index++) {
            writer.WriteElement(new SharedStringItem(new Text("Repeated value " + index.ToString(CultureInfo.InvariantCulture))));
        }

        for (int index = 0; index < 26; index++) {
            writer.WriteElement(new SharedStringItem(new Text("Long segment " + new string((char)('A' + index), 48))));
        }

        for (int row = 1; row <= rowCount; row++) {
            writer.WriteElement(new SharedStringItem(new Text("Distinct value " + row.ToString(CultureInfo.InvariantCulture))));
        }

        writer.WriteEndElement();
    }

    private static void WriteSharedStringFixtureWorksheet(WorksheetPart part, int rowCount) {
        using OpenXmlWriter writer = OpenXmlWriter.Create(part);
        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());
        for (int row = 1; row <= rowCount; row++) {
            uint rowIndex = checked((uint)row);
            writer.WriteStartElement(new Row { RowIndex = rowIndex });
            WriteSharedStringFixtureCell(writer, "A" + row.ToString(CultureInfo.InvariantCulture), row % 12);
            WriteSharedStringFixtureCell(writer, "B" + row.ToString(CultureInfo.InvariantCulture), checked(38 + row - 1));
            WriteSharedStringFixtureCell(writer, "C" + row.ToString(CultureInfo.InvariantCulture), checked(12 + (row % 26)));
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteInlineStringFixtureWorksheet(WorksheetPart part, int rowCount) {
        using OpenXmlWriter writer = OpenXmlWriter.Create(part);
        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());
        for (int row = 1; row <= rowCount; row++) {
            uint rowIndex = checked((uint)row);
            writer.WriteStartElement(new Row { RowIndex = rowIndex });
            WriteInlineStringFixtureCell(writer, "A" + row.ToString(CultureInfo.InvariantCulture), "Repeated value " + (row % 12).ToString(CultureInfo.InvariantCulture));
            WriteInlineStringFixtureCell(writer, "B" + row.ToString(CultureInfo.InvariantCulture), "Distinct value " + row.ToString(CultureInfo.InvariantCulture));
            WriteInlineStringFixtureCell(writer, "C" + row.ToString(CultureInfo.InvariantCulture), "Long segment " + new string((char)('A' + (row % 26)), 48));
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteSharedStringFixtureCell(OpenXmlWriter writer, string reference, int sharedStringIndex) {
        writer.WriteElement(new Cell(new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture))) {
            CellReference = reference,
            DataType = CellValues.SharedString
        });
    }

    private static void WriteInlineStringFixtureCell(OpenXmlWriter writer, string reference, string value) {
        writer.WriteElement(new Cell(new InlineString(new Text(value))) {
            CellReference = reference,
            DataType = CellValues.InlineString
        });
    }

    private static void ValidateStringStorageFixture(
        byte[] workbookBytes,
        int rowCount,
        string expectedCellType,
        int? expectedUniqueSharedStrings) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        ZipArchiveEntry worksheet = archive.GetEntry("xl/worksheets/sheet1.xml")
            ?? throw new InvalidOperationException("The string benchmark fixture has no first worksheet part.");

        int matchingCells = 0;
        using (Stream worksheetStream = worksheet.Open())
        using (XmlReader reader = XmlReader.Create(worksheetStream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit })) {
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.Element
                    && reader.LocalName == "c"
                    && string.Equals(reader.GetAttribute("t"), expectedCellType, StringComparison.Ordinal)) {
                    matchingCells++;
                }
            }
        }

        int expectedCells = checked(rowCount * 3);
        if (matchingCells != expectedCells) {
            throw new InvalidOperationException(
                $"The string benchmark fixture expected {expectedCells.ToString(CultureInfo.InvariantCulture)} cells with t='{expectedCellType}', but found {matchingCells.ToString(CultureInfo.InvariantCulture)}.");
        }

        if (expectedUniqueSharedStrings.HasValue) {
            ZipArchiveEntry sharedStrings = archive.GetEntry("xl/sharedStrings.xml")
                ?? throw new InvalidOperationException("The shared-string benchmark fixture has no shared-string table part.");
            (int declaredCount, int declaredUniqueCount, int itemCount) = ReadSharedStringFixtureCounts(sharedStrings);
            if (declaredCount != expectedCells
                || declaredUniqueCount != expectedUniqueSharedStrings.Value
                || itemCount != expectedUniqueSharedStrings.Value) {
                throw new InvalidOperationException(
                    $"The shared-string benchmark fixture declared count={declaredCount.ToString(CultureInfo.InvariantCulture)}, uniqueCount={declaredUniqueCount.ToString(CultureInfo.InvariantCulture)}, items={itemCount.ToString(CultureInfo.InvariantCulture)}; expected count={expectedCells.ToString(CultureInfo.InvariantCulture)}, uniqueCount/items={expectedUniqueSharedStrings.Value.ToString(CultureInfo.InvariantCulture)}.");
            }
        }
    }

    private static (int DeclaredCount, int DeclaredUniqueCount, int ItemCount) ReadSharedStringFixtureCounts(ZipArchiveEntry entry) {
        int declaredCount = -1;
        int declaredUniqueCount = -1;
        int itemCount = 0;
        using Stream stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        while (reader.Read()) {
            if (reader.NodeType != XmlNodeType.Element) {
                continue;
            }

            if (reader.LocalName == "sst") {
                _ = int.TryParse(reader.GetAttribute("count"), NumberStyles.None, CultureInfo.InvariantCulture, out declaredCount);
                _ = int.TryParse(reader.GetAttribute("uniqueCount"), NumberStyles.None, CultureInfo.InvariantCulture, out declaredUniqueCount);
            } else if (reader.LocalName == "si") {
                itemCount++;
            }
        }

        return (declaredCount, declaredUniqueCount, itemCount);
    }
}
