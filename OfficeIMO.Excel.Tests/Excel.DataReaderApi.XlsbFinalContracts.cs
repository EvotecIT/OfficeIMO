using System.Data.Common;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void XlsbReadersRejectDeclaredGradientStopsWithoutPayload(
        bool useEditableReader) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MissingGradientStops.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceFirstXlsbRecordPayload(
                path,
                "xl/styles.bin",
                recordType: 45,
                data => {
                    byte[] mutated = (byte[])data.Clone();
                    Assert.True(mutated.Length >= 68);
                    WriteUInt32LittleEndian(mutated, 64, 1);
                    return mutated;
                });

            InvalidDataException exception = AssertXlsbReaderRejects(
                path,
                useEditableReader);

            Assert.Contains("BrtFill", exception.Message, StringComparison.Ordinal);
            Assert.Contains("truncated", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(617, false)]
    [InlineData(617, true)]
    [InlineData(626, false)]
    [InlineData(626, true)]
    public void XlsbReadersRejectMissingCustomNumberFormatReferences(
        int collectionBeginType,
        bool useEditableReader) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.MissingNumberFormat.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceFirstXlsbRecordPayloadAfter(
                path,
                "xl/styles.bin",
                collectionBeginType,
                recordType: 47,
                data => {
                    byte[] mutated = (byte[])data.Clone();
                    Assert.True(mutated.Length >= 4);
                    mutated[2] = 164;
                    mutated[3] = 0;
                    return mutated;
                });

            InvalidDataException exception = AssertXlsbReaderRejects(
                path,
                useEditableReader);

            Assert.Contains(
                "missing custom number format 164",
                exception.Message,
                StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("styles", false)]
    [InlineData("styles", true)]
    [InlineData("sharedStrings", false)]
    [InlineData("sharedStrings", true)]
    public void XlsbReadersRejectExternalSingletonMetadataRelationships(
        string relationshipSuffix,
        bool useEditableReader) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.ExternalMetadata.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            MakeWorkbookRelationshipExternal(path, relationshipSuffix);

            InvalidDataException exception = AssertXlsbReaderRejects(
                path,
                useEditableReader);

            Assert.Contains("external", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(
                relationshipSuffix == "styles" ? "styles" : "shared-string",
                exception.Message,
                StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsUnrelatedInternalSheetRelationship() {
        string path = CreateTwoSheetXlsb();
        try {
            ChangeWorksheetRelationshipType(
                path,
                worksheetOccurrence: 1,
                relationshipSuffix: "image");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains(
                "unrelated internal relationship",
                exception.Message,
                StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbSkipsLegitimateNonWorksheetSheetRelationship() {
        string path = CreateTwoSheetXlsb();
        try {
            ChangeWorksheetRelationshipType(
                path,
                worksheetOccurrence: 1,
                relationshipSuffix: "chartsheet");

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbIgnoresStaleWideDimensionForPopulatedSheet() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.StaleWideDimension.{Guid.NewGuid():N}.xlsb");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(2, 2, 42);
                File.WriteAllBytes(path, document.ToBytes(ExcelFileFormat.Xlsb));
            }
            ReplaceXlsbWorksheetLastColumn(path, 16_383);

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(2, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetString(0));
            Assert.Equal(42, reader.GetInt32(1));
        } finally {
            File.Delete(path);
        }
    }

    private static InvalidDataException AssertXlsbReaderRejects(
        string path,
        bool useEditableReader) {
        if (useEditableReader) {
            return Assert.Throws<InvalidDataException>(() => {
                using ExcelDocument document = ExcelDocument.Load(path);
            });
        }

        return Assert.Throws<InvalidDataException>(
            () => ExcelDocument.OpenDataReader(path));
    }

    private static string CreateTwoSheetXlsb() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.SheetRelationships.{Guid.NewGuid():N}.xlsb");
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet first = document.AddWorksheet("First");
        first.CellValue(1, 1, "Value");
        first.CellValue(2, 1, "Ready");
        ExcelSheet second = document.AddWorksheet("Second");
        second.CellValue(1, 1, "Value");
        second.CellValue(2, 1, "Ignored");
        File.WriteAllBytes(path, document.ToBytes(ExcelFileFormat.Xlsb));
        return path;
    }

    private static void MakeWorkbookRelationshipExternal(
        string path,
        string relationshipSuffix) {
        UpdateWorkbookRelationships(path, relationships => {
            XElement relationship = relationships
                .First(element => ((string?)element.Attribute("Type"))
                    ?.EndsWith(
                        "/" + relationshipSuffix,
                        StringComparison.Ordinal) == true);
            relationship.SetAttributeValue(
                "Target",
                $"https://example.invalid/{relationshipSuffix}.bin");
            relationship.SetAttributeValue("TargetMode", "External");
        });
    }

    private static void ChangeWorksheetRelationshipType(
        string path,
        int worksheetOccurrence,
        string relationshipSuffix) {
        UpdateWorkbookRelationships(path, relationships => {
            XElement relationship = relationships
                .Where(element => ((string?)element.Attribute("Type"))
                    ?.EndsWith("/worksheet", StringComparison.Ordinal) == true)
                .ElementAt(worksheetOccurrence);
            string currentType = (string)relationship.Attribute("Type")!;
            int separator = currentType.LastIndexOf('/');
            Assert.True(separator >= 0);
            relationship.SetAttributeValue(
                "Type",
                currentType.Substring(0, separator + 1) + relationshipSuffix);
        });
    }

    private static void UpdateWorkbookRelationships(
        string path,
        Action<IReadOnlyList<XElement>> update) {
        const string entryName = "xl/_rels/workbook.bin.rels";
        byte[] bytes = ReadZipEntry(path, entryName);
        XDocument document;
        using (var input = new MemoryStream(bytes, writable: false)) {
            document = XDocument.Load(input);
        }

        XNamespace relationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        IReadOnlyList<XElement> relationships = document
            .Descendants(relationshipsNamespace + "Relationship")
            .ToArray();
        update(relationships);

        using var output = new MemoryStream();
        document.Save(output);
        ReplaceZipEntry(path, entryName, output.ToArray());
    }

}
