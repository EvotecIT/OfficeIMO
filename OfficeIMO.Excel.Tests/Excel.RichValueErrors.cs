using OfficeIMO.Excel;
using Xunit;
using DocumentFormat.OpenXml.Spreadsheet;
using Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("Blocked", "G1", "#SPILL!")]
        [InlineData("Merged", "G1", "#SPILL!")]
        [InlineData("Table", "G1", "#SPILL!")]
        [InlineData("Edge", "XFD1048576", "#SPILL!")]
        [InlineData("EmptyFilter", "G1", "#CALC!")]
        [InlineData("EmptyUnique", "G1", "#CALC!")]
        [InlineData("ZeroRows", "G1", "#CALC!")]
        [InlineData("ZeroColumns", "G1", "#CALC!")]
        [InlineData("NegativeRows", "G1", "#VALUE!")]
        [InlineData("NegativeColumns", "G1", "#VALUE!")]
        [InlineData("EmptyUnique", "I1", "#CALC!")]
        [InlineData("Blocked", "I1", "#SPILL!")]
        public void RichValueErrors_NativeCachesAgreeAcrossObjectModelAndReaders(string sheetName, string anchor, string expected) {
            string path = Path.Combine(_directoryDocuments, "ExcelFormulaCorpus", "dynamic-spills.xlsx");
            var (row, column) = A1.ParseCellRef(anchor);
            using (var document = ExcelDocument.Load(path, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                var sheet = document[sheetName];
                Assert.Equal(expected, sheet.CellAt(row, column).GetValue().Value);
                Assert.True(sheet.TryGetCellText(row, column, out string text));
                Assert.Equal(expected, text);
                Assert.True(sheet.TryGetCachedFormulaValue(row, column, out string? cached));
                Assert.Equal(expected, cached);
                using var dataReader = document.CreateDataReader(new ExcelReadOptions { SheetName = sheetName, A1Range = anchor + ":" + anchor, HasHeaderRow = false });
                Assert.True(dataReader.Read());
                Assert.Equal(expected, dataReader.GetValue(0));
            }
            using (var reader = ExcelDocumentReader.Open(path)) {
                Assert.Equal(expected, reader.GetSheet(sheetName).ReadRange(anchor + ":" + anchor)[0, 0]);
            }
            var options = new ExcelReadOptions { SheetName = sheetName, A1Range = anchor + ":" + anchor, HasHeaderRow = false };
            using (var reader = ExcelDocument.OpenDataReader(path, options)) {
                Assert.True(reader.Read());
                Assert.Equal(expected, reader.GetValue(0));
            }
            using (var reader = ExcelDocument.OpenDataReader(File.ReadAllBytes(path), options)) {
                Assert.True(reader.Read());
                Assert.Equal(expected, reader.GetValue(0));
            }
        }

        [Theory]
        [InlineData("SEQUENCE(0)", false)]
        [InlineData("SEQUENCE(0)", true)]
        [InlineData("SEQUENCE(1,0)", false)]
        [InlineData("SEQUENCE(1,0)", true)]
        [InlineData("FILTER(A1:A2,B1:B2)", false)]
        [InlineData("FILTER(A1:A2,B1:B2)", true)]
        [InlineData("UNIQUE(A1:A2,FALSE,TRUE)", false)]
        [InlineData("UNIQUE(A1:A2,FALSE,TRUE)", true)]
        public void RichValueErrors_EmptyArraysSaveReopenAndPreserveImages(string formula, bool includeImage) {
            string path = Path.Combine(_directoryWithFiles, "ModernErrors.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, 1d); sheet.CellValue(2, 1, 1d);
                sheet.CellValue(1, 2, 0d); sheet.CellValue(2, 2, 0d);
                if (includeImage) sheet.SetInCellImage(1, 4, TinyPng, altText: "Retained image");
                sheet.SetArrayFormula("G1:G1", formula);
                sheet.CellFormula(1, 9, "G1");
                sheet.CellFormula(1, 10, "IFERROR(G1,\"empty\")");
                Assert.Equal(3, document.Calculate());
                Assert.Equal("#CALC!", sheet.CellAt(1, 7).GetValue().Value);
                Assert.Equal("#CALC!", sheet.CellAt(1, 9).GetValue().Value);
                Assert.Equal("empty", sheet.CellAt(1, 10).GetValue().Value);
                int valueCount = document.WorkbookPartRoot.RdRichValueParts.Single().RichValueData!.ChildElements.Count;
                Assert.Equal(3, document.Calculate());
                Assert.Equal(valueCount, document.WorkbookPartRoot.RdRichValueParts.Single().RichValueData!.ChildElements.Count);
                if (includeImage) Assert.Equal(TinyPng, Assert.Single(sheet.GetInCellImages()).Bytes);
                document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
            }
            using (var document = ExcelDocument.Load(path)) {
                Assert.Equal("#CALC!", document["Data"].CellAt(1, 7).GetValue().Value);
                if (includeImage) Assert.Equal(TinyPng, Assert.Single(document["Data"].GetInCellImages()).Bytes);
                Assert.Empty(document.ValidateOpenXml());
            }
            AssertWorkbookOpensViaExcelComWhenAvailable(path, "Modern cached errors and retained images must open and calculate in desktop Excel.",
                new Dictionary<string, string> { ["G1"] = "#CALC!", ["I1"] = "#CALC!", ["J1"] = "empty" });
        }

        [Theory]
        [InlineData("SORT(UNIQUE(A1:A2))", "_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A2))")]
        [InlineData("FILTER(A1:A2,B1:B2,\"SEQUENCE(0)\")", "_xlfn._xlws.FILTER(A1:A2,B1:B2,\"SEQUENCE(0)\")")]
        [InlineData("_xlfn.SEQUENCE(0)", "_xlfn.SEQUENCE(0)")]
        [InlineData("SUM('SEQUENCE(0)'!A1)", "SUM('SEQUENCE(0)'!A1)")]
        [InlineData("SUM(Table1[SEQUENCE(0)])", "SUM(Table1[SEQUENCE(0)])")]
        public void ArrayFormulaAuthoring_QualifiesFunctionsAndPreservesReferenceText(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.SetArrayFormula("G1:G1", formula);
            Assert.Equal(expected, sheet.GetFormulaText(1, 7));
        }

        [Theory]
        [InlineData("sheet")]
        [InlineData("document")]
        [InlineData("save")]
        [InlineData("array")]
        public void RichValueErrors_CacheClearingDoesNotResurrectErrors(string mode) {
            string path = Path.Combine(_directoryWithFiles, "ClearedErrors.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.SetInCellImage(1, 4, TinyPng, altText: "Retained image");
                sheet.SetArrayFormula("G1:G1", "SEQUENCE(0)");
                Assert.Equal(1, document.Calculate());
                Assert.Equal("#CALC!", sheet.CellAt(1, 7).GetValue().Value);
                if (mode == "sheet") sheet.ClearCachedFormulaResults();
                else if (mode == "document") document.ClearCachedFormulaResults();
                else if (mode == "array") sheet.ClearArrayFormula("G1");
                document.Save(new ExcelSaveOptions { ClearCachedFormulaResultsBeforeSave = mode == "save" });
                sheet = document["Data"];
                Assert.False(sheet.TryGetCachedFormulaValue(1, 7, out _));
                Assert.Null(sheet.CellAt(1, 7).GetValue().Value);
                Assert.Null(sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "G1").ValueMetaIndex);
                Assert.Equal(TinyPng, Assert.Single(sheet.GetInCellImages()).Bytes);
            }
            using (var document = ExcelDocument.Load(path)) {
                Assert.False(document["Data"].TryGetCachedFormulaValue(1, 7, out _));
                Assert.Equal(TinyPng, Assert.Single(document["Data"].GetInCellImages()).Bytes);
            }
            using var reader = ExcelDocumentReader.Open(path);
            Assert.Equal(mode == "array" ? null : "_xlfn.SEQUENCE(0)", reader.GetSheet("Data").ReadRange("G1:G1")[0, 0]);
        }

        [Fact]
        public void RichValueErrors_MissingSerializedPayloadDoesNotCreateACache() {
            string path = Path.Combine(_directoryWithFiles, "MissingPayload.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.SetArrayFormula("G1:G1", "SEQUENCE(0)");
                Assert.Equal(1, document.Calculate());
                sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single().CellValue = null;
                Assert.False(sheet.TryGetCachedFormulaValue(1, 7, out _));
                document.Save();
            }
            using var reader = ExcelDocument.OpenDataReader(path,
                new ExcelReadOptions { SheetName = "Data", A1Range = "G1:G1", HasHeaderRow = false });
            Assert.True(reader.Read());
            Assert.Equal("_xlfn.SEQUENCE(0)", reader.GetValue(0));
        }

        [Fact]
        public void RichValueErrors_DependencyGuardInvalidatesRichCache() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.SetArrayFormula("G1:G1", "SEQUENCE(0)");
            Assert.Equal(1, document.Calculate());
            Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single();
            cell.CellFormula = new CellFormula("G1");
            Assert.Equal(0, document.Calculate());
            Assert.False(sheet.TryGetCachedFormulaValue(1, 7, out _));
            Assert.Null(cell.ValueMetaIndex);
            Assert.True(cell.CellFormula.CalculateCell!.Value);
        }

        [Theory]
        [InlineData("xl/metadata.xml")]
        [InlineData("xl/richData/rdrichvalue.xml")]
        [InlineData("xl/richData/rdrichvaluestructure.xml")]
        public void RichValueErrors_OversizedPartsAreRejectedBeforeDomMaterialization(string entryName) {
            string path = Path.Combine(_directoryWithFiles, "OversizedErrors.xlsx");
            File.Copy(Path.Combine(_directoryDocuments, "ExcelFormulaCorpus", "dynamic-spills.xlsx"), path);
            using (var archive = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update)) {
                var entry = archive.GetEntry(entryName)!;
                string xml;
                using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
                entry.Delete();
                using var writer = new StreamWriter(archive.CreateEntry(entryName).Open());
                int rootStart = xml.IndexOf('>', xml.IndexOf("?>", StringComparison.Ordinal) + 2) + 1;
                writer.Write(xml.Substring(0, rootStart));
                string padding = new string(' ', 1024);
                for (int index = 0; index < 16 * 1024 + 1; index++) writer.Write(padding);
                writer.Write(xml.Substring(rootStart));
            }
            using var document = ExcelDocument.Load(path, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            var parts = new DocumentFormat.OpenXml.Packaging.OpenXmlPart[] {
                document.WorkbookPartRoot.CellMetadataPart!,
                document.WorkbookPartRoot.RdRichValueParts.Single(),
                document.WorkbookPartRoot.GetPartsOfType<DocumentFormat.OpenXml.Packaging.RdRichValueStructurePart>().Single()
            };
            Assert.All(parts, part => Assert.False(part.IsRootElementLoaded));
            Assert.Throws<InvalidDataException>(() => document["EmptyUnique"].CellAt(1, 7).GetValue());
            Assert.All(parts, part => Assert.False(part.IsRootElementLoaded));
        }

        [Fact]
        public void RichValueErrors_EditableRootsDoNotKeepStaleErrorMappings() {
            string path = Path.Combine(_directoryWithFiles, "Editable.xlsx");
            File.Copy(Path.Combine(_directoryDocuments, "ExcelFormulaCorpus", "dynamic-spills.xlsx"), path);
            using var document = ExcelDocument.Load(path);
            var sheet = document["EmptyUnique"];
            Assert.Equal("#CALC!", sheet.CellAt(1, 7).GetValue().Value);
            Rich.RichValueData values = document.WorkbookPartRoot.RdRichValueParts.Single().RichValueData!;
            values.Elements<Rich.RichValue>().First().Elements<Rich.Value>().First().Text = "8";
            Assert.Equal("#SPILL!", sheet.CellAt(1, 7).GetValue().Value);
        }

        [Fact]
        public void RichValueErrors_UnresolvableMetadataRetainsOrdinaryFallback() {
            string path = Path.Combine(_directoryWithFiles, "InvalidMetadata.xlsx");
            File.Copy(Path.Combine(_directoryDocuments, "ExcelFormulaCorpus", "dynamic-spills.xlsx"), path);
            using var document = ExcelDocument.Load(path);
            var sheet = document["EmptyUnique"];
            Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == "G1");
            cell.ValueMetaIndex = uint.MaxValue;
            Assert.Equal("#VALUE!", sheet.CellAt(1, 7).GetValue().Value);
            cell.ValueMetaIndex = 1;
            var metadata = document.WorkbookPartRoot.CellMetadataPart!.Metadata!;
            metadata.GetFirstChild<ValueMetadata>()!.Elements<MetadataBlock>().First().Elements<MetadataRecord>().First().RemoveAttribute("v", "");
            Assert.Equal("#VALUE!", sheet.CellAt(1, 7).GetValue().Value);
            metadata.GetFirstChild<ValueMetadata>()!.Elements<MetadataBlock>().First().Elements<MetadataRecord>().First().Val = 0;
            document.WorkbookPartRoot.GetPartsOfType<DocumentFormat.OpenXml.Packaging.RdRichValueStructurePart>()
                .Single().RichValueStructures!.Elements<Rich.RichValueStructure>().First()
                .Elements<Rich.Key>().Single(key => key.N?.Value == "errorType").RemoveAttribute("t", "");
            Assert.Equal("#VALUE!", sheet.CellAt(1, 7).GetValue().Value);
        }
    }
}
