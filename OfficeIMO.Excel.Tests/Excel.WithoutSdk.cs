using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// A plain workbook loads, reads, takes plain cell values and saves without opening the Open XML SDK package. The
    /// owned save must give the cells the SDK model's save gives them, and leave the other parts as they were read.
    /// </summary>
    public class ExcelWithoutSdk {
        private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        [Fact]
        public void LoadReadEditAndSaveDoNotOpenTheSdk() {
            byte[] source = CreateWorkbook();
            using var document = ExcelDocument.Load(new MemoryStream(source));
            Assert.Equal(new[] { "Data", "Second" }, document.Sheets.Select(sheet => sheet.Name));
            int cells = document.Sheets[0].EnumerateCells().Count();
            Assert.Equal(12, cells);

            ApplyEdits(document);
            using var output = new MemoryStream();
            document.Save(output);

            Assert.Empty(document.OpenXmlSdkFallbacks);
        }

        [Fact]
        public void SavedCellsMatchASaveThroughTheSdk() {
            byte[] source = CreateWorkbook();
            Assert.Equal(CellsOf(SaveThroughSdk(source, ApplyEdits)), CellsOf(SaveWithoutSdk(source, ApplyEdits, out _)));
        }

        [Fact]
        public void PartsTheEditsDoNotTouchAreCopiedAsStored() {
            byte[] source = CreateWorkbook();
            byte[] saved = SaveWithoutSdk(source, ApplyEdits, out _);
            var changed = ChangedEntries(source, saved);
            Assert.Equal(new[] { "docProps/core.xml", "xl/sharedStrings.xml", "xl/styles.xml", "xl/worksheets/sheet1.xml" }, changed.OrderBy(name => name, StringComparer.Ordinal));
        }

        [Fact]
        public void AnOperationTheOwnedPathDoesNotCoverOpensTheSdkWithTheValuesWritten() {
            byte[] source = CreateWorkbook();
            byte[] expected = SaveThroughSdk(source, document => {
                ApplyEdits(document);
                document.Sheets[0].Cell(6, 6).SetFormula("=1+1");
            });
            byte[] actual;
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                ApplyEdits(document);
                document.Sheets[0].Cell(6, 6).SetFormula("=1+1");
                Assert.Single(document.OpenXmlSdkFallbacks);
                using var output = new MemoryStream();
                document.Save(output);
                actual = output.ToArray();
            }
            Assert.Equal(CellsOf(expected), CellsOf(actual));
        }

        [Fact]
        public void EditsThroughTheSdkPackageMixWithEditsThroughTheModel() {
            byte[] source = CreateWorkbook();
            byte[] saved;
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                document.Sheets[0].Cell(2, 1).SetValue("before the package");
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument package = document.OpenXmlDocument;
                AppendInlineText(package, "Second", 5, "through the package");
                document.Sheets[0].Cell(3, 1).SetValue("after the package");
                document.Sheets[1].Cell(6, 1).SetValue(9.5);
                Assert.Same(package, document.OpenXmlDocument);
                using var output = new MemoryStream();
                document.Save(output);
                saved = output.ToArray();
            }

            using var reread = ExcelDocument.Load(new MemoryStream(saved));
            Assert.Equal("before the package", reread.Sheets[0].Cell(2, 1).GetValue().Value);
            Assert.Equal("after the package", reread.Sheets[0].Cell(3, 1).GetValue().Value);
            Assert.Equal("through the package", reread.Sheets[1].Cell(5, 1).GetValue().Value);
            Assert.Equal(9.5, reread.Sheets[1].Cell(6, 1).GetValue().Value);
            Assert.Equal("Untouched", reread.Sheets[1].Cell(1, 1).GetValue().Value);
        }

        [Fact]
        public void AWorkbookCreatedAndEditedOnlyThroughTheSdkPackageSaves() {
            byte[] saved;
            using (var stream = new MemoryStream()) {
                using (var document = ExcelDocument.Create(stream)) {
                    DocumentFormat.OpenXml.Packaging.SpreadsheetDocument package = document.OpenXmlDocument;
                    var workbookPart = package.WorkbookPart!;
                    var worksheetPart = workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
                    worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
                    var sheets = workbookPart.Workbook.Sheets ??= workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
                    sheets.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet {
                        Name = "Raw",
                        SheetId = 1U,
                        Id = workbookPart.GetIdOfPart(worksheetPart)
                    });
                    AppendInlineText(package, "Raw", 1, "only the package");
                    document.Save();
                }
                saved = stream.ToArray();
            }

            using var reread = ExcelDocument.Load(new MemoryStream(saved));
            Assert.Equal("Raw", reread.Sheets[0].Name);
            Assert.Equal("only the package", reread.Sheets[0].Cell(1, 1).GetValue().Value);
        }

        private static void AppendInlineText(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument package, string sheetName, int row, string text) {
            var workbookPart = package.WorkbookPart!;
            var sheet = workbookPart.Workbook.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Single(candidate => candidate.Name == sheetName);
            var worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()!;
            sheetData.Append(new DocumentFormat.OpenXml.Spreadsheet.Row(new DocumentFormat.OpenXml.Spreadsheet.Cell(
                new DocumentFormat.OpenXml.Spreadsheet.InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(text))) {
                CellReference = "A" + row.ToString(CultureInfo.InvariantCulture),
                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString
            }) { RowIndex = (uint)row });
        }

        [Fact]
        public void ASheetAddedToALoadedWorkbookSavesWithoutTheSdk() {
            byte[] source = CreateWorkbook();
            void AddAndFill(ExcelDocument document) {
                ExcelSheet added = document.AddSheet("Added");
                added.Cell(1, 1).SetValue("Name").SetBold();
                added.Cell(2, 1).SetValue(12.5);
                document.Sheets[0].Cell(2, 1).SetValue("Changed");
            }
            byte[] owned = SaveWithoutSdk(source, AddAndFill, out bool stayedOwned);
            Assert.True(stayedOwned);
            Assert.Equal(CellsOf(SaveThroughSdk(source, AddAndFill)), CellsOf(owned));
            using var reread = ExcelDocument.Load(new MemoryStream(owned));
            Assert.Equal(new[] { "Data", "Second", "Added" }, reread.Sheets.Select(sheet => sheet.Name));
        }

        [Fact]
        public void AWorkbookCreatedFromATemplateTakesASheetAndSavesWithoutTheSdk() {
            string folder = Path.Combine(Path.GetTempPath(), "officeimo-template-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            try {
                string template = Path.Combine(folder, "template.xlsx");
                File.WriteAllBytes(template, CreateWorkbook());
                string output = Path.Combine(folder, "output.xlsx");
                using (var document = ExcelDocument.Create(output, new ExcelCreateOptions { Template = template })) {
                    ExcelSheet added = document.GetOrCreateSheet("Added", ExcelSheetNameValidationMode.Sanitize);
                    added.Cell(1, 1).SetValue(100);
                    Assert.True(document.HasWorksheets);
                    document.Save(output);
                    Assert.Empty(document.OpenXmlSdkFallbacks);
                }

                using var reread = ExcelDocument.Load(output);
                Assert.Equal(new[] { "Data", "Second", "Added" }, reread.Sheets.Select(sheet => sheet.Name));
                Assert.Equal(100d, reread.Sheets[2].Cell(1, 1).GetValue().Value);
            } finally {
                Directory.Delete(folder, recursive: true);
            }
        }

        [Fact]
        public void TheFirstSdkReadInsideAReadScopeKeepsTheRecordedChanges() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            document.Sheets[0].Cell(2, 1).SetValue("Changed");
            int sheets = document.ExecuteReadWithoutMaterializing(() => document.OpenXmlDocument.WorkbookPart!.Workbook.Sheets!.Count());

            Assert.Equal(2, sheets);
            Assert.Single(document.OpenXmlSdkFallbacks);
            Assert.Equal("Changed", document.Sheets[0].Cell(2, 1).GetValue().Value);
        }

        [Theory]
        [InlineData("<s:sheetData xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>")]
        [InlineData("<sheetData><p:row xmlns:p=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" r=\"2\"><p:c r=\"A2\"><p:v>7</p:v></p:c></p:row></sheetData>")]
        public void RowsAddedBesideLocallyDeclaredPrefixesStayWellFormed(string sheetData) {
            byte[] source = WithEntry(CreateWorkbook(), "xl/worksheets/sheet1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + sheetData + "</worksheet>");
            byte[] saved = SaveWithoutSdk(source, document => document.Sheets[0].Cell(1, 1).SetValue(5), out bool owned);

            Assert.True(owned);
            using var reread = ExcelDocument.Load(new MemoryStream(saved));
            Assert.Equal(5d, reread.Sheets[0].Cell(1, 1).GetValue().Value);
        }

        [Fact]
        public void SignatureMetadataInTheApplicationPropertiesKeepsTheSdkLoad() {
            byte[] source = WithEntry(CreateWorkbook(), "docProps/app.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
                + "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><DigSig><vt:blob>AA==</vt:blob></DigSig></Properties>");
            using var document = ExcelDocument.Load(new MemoryStream(source));
            Assert.False(document.IsSdkDeferred);
        }

        [Fact]
        public void DisablingTheFastPackageWriterSavesThroughTheModel() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            document.Sheets[0].Cell(2, 1).SetValue("Changed");
            using var output = new MemoryStream();
            document.Save(output, new ExcelSaveOptions { DisableFastPackageWriter = true });

            Assert.Single(document.OpenXmlSdkFallbacks);
        }

        [Fact]
        public async Task AsyncSavesUseTheOwnedPath() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            document.Sheets[0].Cell(2, 1).SetValue("Changed");
            using var output = new MemoryStream();
            await document.SaveAsync(output);

            Assert.Empty(document.OpenXmlSdkFallbacks);
            using var reread = ExcelDocument.Load(new MemoryStream(output.ToArray()));
            Assert.Equal("Changed", reread.Sheets[0].Cell(2, 1).GetValue().Value);
        }

        [Fact]
        public void AnExistingRichTextItemIsReusedAndNewItemsGoBeforeTheExtensionList() {
            byte[] source = CreateWorkbook();
            XDocument table = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(source, "xl/sharedStrings.xml")));
            int existing = table.Root!.Elements(Main + "si").Count();
            table.Root.Add(new XElement(Main + "si", new XElement(Main + "r", new XElement(Main + "t", "Alpha"))),
                new XElement(Main + "extLst", new XElement(Main + "ext", new XAttribute("uri", "{00000000-0000-0000-0000-000000000001}"))));
            source = WithEntry(source, "xl/sharedStrings.xml", table.ToString(SaveOptions.DisableFormatting));

            byte[] saved = SaveWithoutSdk(source, document => {
                document.Sheets[0].Cell(6, 1).SetValue("Alpha");
                document.Sheets[0].Cell(7, 1).SetValue("Beta");
            }, out bool owned);

            Assert.True(owned);
            XElement root = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(saved, "xl/sharedStrings.xml"))).Root!;
            Assert.Equal(existing + 2, root.Elements(Main + "si").Count());
            Assert.Equal(Main + "extLst", root.Elements().Last().Name);
            XElement sheet = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(saved, "xl/worksheets/sheet1.xml"))).Root!;
            XElement alpha = sheet.Descendants(Main + "c").Single(cell => (string?)cell.Attribute("r") == "A6");
            Assert.Equal(existing.ToString(CultureInfo.InvariantCulture), alpha.Element(Main + "v")!.Value);
        }

        [Fact]
        public void AnItemWhoseTextDependsOnCompatibilityMarkupIsNotReused() {
            byte[] source = CreateWorkbook();
            XDocument table = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(source, "xl/sharedStrings.xml")));
            int existing = table.Root!.Elements(Main + "si").Count();
            XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            table.Root.Add(new XElement(Main + "si", new XAttribute(XNamespace.Xmlns + "mc", mc.NamespaceName),
                new XElement(Main + "r", new XElement(Main + "t", "Alpha")),
                new XElement(mc + "AlternateContent", new XElement(mc + "Fallback", new XElement(Main + "r", new XElement(Main + "t", "Beta"))))));
            source = WithEntry(source, "xl/sharedStrings.xml", table.ToString(SaveOptions.DisableFormatting));

            byte[] saved = SaveWithoutSdk(source, document => document.Sheets[0].Cell(6, 1).SetValue("Alpha"), out bool owned);

            Assert.True(owned);
            XElement sheet = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(saved, "xl/worksheets/sheet1.xml"))).Root!;
            XElement alpha = sheet.Descendants(Main + "c").Single(cell => (string?)cell.Attribute("r") == "A6");
            Assert.Equal((existing + 1).ToString(CultureInfo.InvariantCulture), alpha.Element(Main + "v")!.Value);
        }

        [Fact]
        public void ACellReferenceWrittenAsACharacterReferenceKeepsItsColumn() {
            byte[] source = WithEntry(CreateWorkbook(), "xl/worksheets/sheet1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<sheetData><row r=\"99\"><c r=\"&#90;99\"><v>7</v></c></row></sheetData></worksheet>");
            byte[] saved = SaveWithoutSdk(source, document => document.Sheets[0].Cell(1, 1).SetValue(5), out _);

            using var reread = ExcelDocument.Load(new MemoryStream(saved));
            Assert.Equal(5d, reread.Sheets[0].Cell(1, 1).GetValue().Value);
            Assert.Equal(7d, reread.Sheets[0].Cell(99, 26).GetValue().Value);
            XElement sheet = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(saved, "xl/worksheets/sheet1.xml"))).Root!;
            string? dimension = (string?)sheet.Element(Main + "dimension")?.Attribute("ref");
            Assert.True(dimension == null || dimension == "A1:Z99", dimension);
        }

        [Fact]
        public void ABorderWithAnEdgeInCompatibilityMarkupIsChangedThroughTheModel() {
            byte[] source = CreateWorkbook();
            XDocument styles = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(source, "xl/styles.xml")));
            XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            // Declared on the root, so the wrapper is kept inside the border rather than refused with the stylesheet.
            styles.Root!.SetAttributeValue(XNamespace.Xmlns + "mc", mc.NamespaceName);
            XElement border = styles.Root.Element(Main + "borders")!.Elements(Main + "border").First();
            border.AddFirst(new XElement(mc + "AlternateContent",
                new XElement(mc + "Fallback", new XElement(Main + "left", new XAttribute("style", "thick")))));
            source = WithEntry(source, "xl/styles.xml", styles.ToString(SaveOptions.DisableFormatting));

            byte[] saved = SaveWithoutSdk(source, document => document.Sheets[0].Cell(1, 1).SetBorder(ExcelBorderStyle.Thin), out bool owned);

            Assert.False(owned);
            using var reread = ExcelDocument.Load(new MemoryStream(saved));
            Assert.Equal("Name", reread.Sheets[0].Cell(1, 1).GetValue().Value);
        }

        [Fact]
        public void ASaveOverTheInMemoryLimitWritesNothingAndKeepsTheChanges() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            document.Sheets[0].Cell(2, 1).SetValue("Changed");
            using var output = new MemoryStream();

            Assert.Throws<InvalidDataException>(() => document.Save(output, new ExcelSaveOptions { MaxInMemoryPackageBytes = 1024 }));

            Assert.Equal(0, output.Length);
            Assert.Empty(document.OpenXmlSdkFallbacks);
            document.Save(output);
            Assert.Empty(document.OpenXmlSdkFallbacks);
            using var reread = ExcelDocument.Load(new MemoryStream(output.ToArray()));
            Assert.Equal("Changed", reread.Sheets[0].Cell(2, 1).GetValue().Value);
        }

        [Fact]
        public void AFailedDestinationDoesNotReportTheOwnedSave() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            document.Sheets[0].Cell(2, 1).SetValue("Changed");
            var output = new CallbackStream { OnWrite = () => throw new IOException("The destination is full.") };

            Assert.Throws<IOException>(() => document.Save(output));

            Assert.NotEqual(ExcelSavePackageWriter.OwnedPackage, document.LastSaveDiagnostics.Writer);
            using var retry = new MemoryStream();
            document.Save(retry);
            Assert.Equal(ExcelSavePackageWriter.OwnedPackage, document.LastSaveDiagnostics.Writer);
        }

        [Fact]
        public void TheLastCustomNumberFormatIdIsNeverFollowedByGeneral() {
            byte[] source = CreateWorkbook();
            XDocument styles = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(source, "xl/styles.xml")));
            XElement root = styles.Root!;
            XElement? formats = root.Element(Main + "numFmts");
            if (formats == null) {
                formats = new XElement(Main + "numFmts");
                root.AddFirst(formats);
            }
            formats.Add(new XElement(Main + "numFmt", new XAttribute("numFmtId", "4294967294"), new XAttribute("formatCode", "0.000")));
            formats.SetAttributeValue("count", formats.Elements(Main + "numFmt").Count());
            source = WithEntry(source, "xl/styles.xml", styles.ToString(SaveOptions.DisableFormatting));

            byte[]? saved = null;
            try {
                saved = SaveWithoutSdk(source, document => {
                    document.Sheets[0].Cell(5, 1).SetValue(1).SetNumberFormat("0.0000");
                    document.Sheets[0].Cell(6, 1).SetValue(2).SetNumberFormat("0.00000");
                }, out _);
            } catch (OverflowException) {
                // No ID is left for the second format; refusing the save is a valid outcome.
                return;
            }

            XElement savedStyles = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(saved, "xl/styles.xml"))).Root!;
            Assert.DoesNotContain(savedStyles.Descendants(Main + "numFmt"), format => (string?)format.Attribute("numFmtId") == "0");
        }

        [Fact]
        public void ReadsAfterRecordedChangesSeeEveryRecordedCell() {
            // An empty sheet takes the replayed values into the cell buffer once the SDK opens.
            using var document = ExcelDocument.Load(new MemoryStream(CreateEmptySheetWorkbook()));
            ExcelSheet sheet = document.Sheets[0];
            int before = sheet.EnumerateCells().Count();
            for (int row = 10; row < 110; row++) {
                sheet.Cell(row, 1).SetValue("Name " + row);
                sheet.Cell(row, 2).SetValue(row * 1.5);
                sheet.Cell(row, 3).SetValue(new DateTime(2024, 1, 1).AddDays(row));
            }

            Assert.Equal(before + 300, sheet.EnumerateCells().Count());
            Assert.Equal("A109", sheet.FindFirst("Name 109"));
        }

        [Fact]
        public void ANewWorkbookWithAThemeIsCreatedWithoutTheSdk() {
            OfficeTheme theme = OfficeTheme.Classic;
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Create(output, new ExcelCreateOptions { Theme = theme })) {
                document.AddSheet("Data").Cell(1, 1).SetValue("Themed");
                document.Save();
                Assert.Empty(document.OpenXmlSdkFallbacks);
            }

            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XElement themeRoot = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(output.ToArray(), "xl/theme/theme1.xml"))).Root!;
            Assert.Equal(theme.Name, (string?)themeRoot.Attribute("name"));
            Assert.Equal(theme.Fonts.Body.Family, (string?)themeRoot.Descendants(drawing + "minorFont").Single().Element(drawing + "latin")!.Attribute("typeface"));
            XElement font = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(output.ToArray(), "xl/styles.xml"))).Root!
                .Element(Main + "fonts")!.Elements(Main + "font").First();
            Assert.Equal(theme.Fonts.Body.Family, (string?)font.Element(Main + "name")!.Attribute("val"));
            Assert.Equal(theme.TypeScale.Body.ToPoints(), double.Parse((string)font.Element(Main + "sz")!.Attribute("val")!, CultureInfo.InvariantCulture));
        }

        [Fact]
        public void AThemeAppliedToANewWorkbookSurvivesTheTabularWriter() {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Create(output)) {
                document.ApplyTheme(OfficeTheme.Dark);
                ExcelSheet sheet = document.AddSheet("Data");
                for (int row = 1; row <= 200; row++) sheet.Cell(row, 1).SetValue(row * 1.5);
                document.Save();
                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            }

            XElement themeRoot = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(output.ToArray(), "xl/theme/theme1.xml"))).Root!;
            Assert.Equal(OfficeTheme.Dark.Name, (string?)themeRoot.Attribute("name"));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ANewWorkbookSavedOnDisposeWritesItsDestinationWithoutTheSdk(bool toFile) {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            using var destination = new MemoryStream();
            var options = new ExcelCreateOptions { PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose };
            IReadOnlyList<string> fallbacks;
            using (ExcelDocument document = toFile ? ExcelDocument.Create(path, options) : ExcelDocument.Create(destination, options)) {
                document.AddSheet("Data").Cell(1, 1).SetValue("Saved on dispose");
                fallbacks = document.OpenXmlSdkFallbacks;
                Assert.Empty(fallbacks);
            }

            try {
                byte[] saved = toFile ? File.ReadAllBytes(path) : destination.ToArray();
                using var reread = ExcelDocument.Load(new MemoryStream(saved));
                Assert.Equal("Saved on dispose", reread.Sheets[0].Cell(1, 1).GetValue().Value);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void AWorkbookCreatedFromATemplateWithAThemeStaysOffTheSdk() {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Create(output, new ExcelCreateOptions { TemplateStream = new MemoryStream(CreateWorkbook()), Theme = OfficeTheme.Dark })) {
                Assert.True(document.IsSdkDeferred);
                document.Sheets[0].Cell(10, 1).SetValue("From template");
                document.Save();
                Assert.Empty(document.OpenXmlSdkFallbacks);
            }

            XElement themeRoot = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(output.ToArray(), "xl/theme/theme1.xml"))).Root!;
            Assert.Equal(OfficeTheme.Dark.Name, (string?)themeRoot.Attribute("name"));
            using var reread = ExcelDocument.Load(new MemoryStream(output.ToArray()));
            Assert.Equal("From template", reread.Sheets[0].Cell(10, 1).GetValue().Value);
        }

        [Theory]
        [InlineData(".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.main+xml")]
        [InlineData(".xltx", "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml")]
        [InlineData(".xlam", "application/vnd.ms-excel.addin.macroEnabled.main+xml")]
        public void ANewWorkbookTakesTheTypeOfItsFileWithoutTheSdk(string extension, string contentType) {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
            try {
                using (var document = ExcelDocument.Create(path)) {
                    // Enough plain values for the tabular writer, which declares an XLSX workbook, so the type must survive it.
                    ExcelSheet sheet = document.AddSheet("Data");
                    for (int row = 1; row <= 200; row++) sheet.Cell(row, 1).SetValue(row);
                    document.Save();
                    Assert.Empty(document.OpenXmlSdkFallbacks);
                }
                XElement types = XDocument.Parse(Encoding.UTF8.GetString(EntryBytes(File.ReadAllBytes(path), "[Content_Types].xml"))).Root!;
                XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                Assert.Contains(types.Elements(ct + "Override"), element => (string?)element.Attribute("PartName") == "/xl/workbook.xml"
                    && (string?)element.Attribute("ContentType") == contentType);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void ANewWorkbookSavesTheSameCellsAsOneBuiltThroughTheSdk() {
            static byte[] Build(bool throughSdk) {
                using var output = new MemoryStream();
                using (var document = ExcelDocument.Create(output)) {
                    if (throughSdk) document.OpenDeferredSdk("test");
                    foreach (string name in new[] { "First", "Second" }) {
                        ExcelSheet sheet = document.AddSheet(name);
                        sheet.Cell(1, 1).SetValue("Header");
                        sheet.Cell(1, 1).SetBold();
                        for (int row = 2; row <= 300; row++) {
                            sheet.Cell(row, 1).SetValue("Item " + row);
                            sheet.Cell(row, 2).SetValue(row * 1.5);
                            sheet.Cell(row, 3).SetValue(new DateTime(2024, 1, 1).AddDays(row));
                            sheet.Cell(row, 4).SetValue(row % 2 == 0);
                            sheet.Cell(row, 5).SetValue(row * 0.01m);
                        }
                        sheet.Cell(2, 2).SetNumberFormat("#,##0.00");
                    }
                    document.Save();
                }
                return output.ToArray();
            }

            Assert.Equal(CellsOf(Build(throughSdk: true)), CellsOf(Build(throughSdk: false)));
        }

        [Fact]
        public void AStyleOnANewWorkbookMovesItsBufferedValuesToTheModelInOrder() {
            static byte[] Build(bool throughSdk) {
                using var output = new MemoryStream();
                using (var document = ExcelDocument.Create(output)) {
                    if (throughSdk) document.OpenDeferredSdk("test");
                    ExcelSheet sheet = document.AddSheet("Data");
                    for (int row = 1; row <= 200; row++) {
                        sheet.Cell(row, 1).SetValue("Item " + row);
                        sheet.Cell(row, 2).SetValue(row * 0.5m);
                    }
                    sheet.Cell(5, 1).SetBold();
                    sheet.Cell(3, 2).SetValue("Replaced");
                    sheet.Cell(201, 1).SetValue(new DateTime(2024, 5, 6));
                    document.Save();
                    // The style moves the new workbook to the model once; the values written before it come along.
                    if (!throughSdk) Assert.Single(document.OpenXmlSdkFallbacks);
                }
                return output.ToArray();
            }
            // Either route may store text as shared or inline strings; the values and the style are what must match.
            static List<string> Read(byte[] package) {
                using var document = ExcelDocument.Load(new MemoryStream(package));
                ExcelSheet sheet = document.Sheets[0];
                var cells = sheet.EnumerateCells().Select(cell => cell.Row + "," + cell.Column + "=" + Convert.ToString(cell.Value, CultureInfo.InvariantCulture)).ToList();
                cells.Add("A5 bold=" + sheet.Cell(5, 1).GetStyle().Bold);
                return cells;
            }

            Assert.Equal(Read(Build(throughSdk: true)), Read(Build(throughSdk: false)));
        }

        [Fact]
        public void ANewSheetOfValuesSavesThroughTheTabularWriterWithoutTheSdk() {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Create(output)) {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int row = 1; row <= 200; row++) {
                    sheet.Cell(row, 1).SetValue(row * 1.25);
                    sheet.Cell(row, 2).SetValue("Name " + row);
                }
                document.Save();
                Assert.Empty(document.OpenXmlSdkFallbacks);
                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                // The package the tabular writer produced is the workbook's state, so later edits and saves work on it.
                sheet.Cell(201, 1).SetValue("After");
                document.Save();
            }

            using var reread = ExcelDocument.Load(new MemoryStream(output.ToArray()));
            Assert.Equal(250d, reread.Sheets[0].Cell(200, 1).GetValue().Value);
            Assert.Equal("After", reread.Sheets[0].Cell(201, 1).GetValue().Value);
        }

        [Fact]
        public async Task ACanceledAsyncSaveWritesNothing() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            document.Sheets[0].Cell(2, 1).SetValue("Changed");
            using var output = new MemoryStream();
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.SaveAsync(output, cancellationToken: canceled.Token));

            Assert.Equal(0, output.Length);
            await document.SaveAsync(output);
            using var reread = ExcelDocument.Load(new MemoryStream(output.ToArray()));
            Assert.Equal("Changed", reread.Sheets[0].Cell(2, 1).GetValue().Value);
        }

        [Fact]
        public async Task AsyncSavesWriteToDestinationsThatOnlyWriteAsynchronously() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            document.Sheets[0].Cell(2, 1).SetValue("Changed");
            var output = new CallbackStream { RejectSynchronousWrites = true };
            await document.SaveAsync(output);

            Assert.Empty(document.OpenXmlSdkFallbacks);
            using var reread = ExcelDocument.Load(new MemoryStream(output.Inner.ToArray()));
            Assert.Equal("Changed", reread.Sheets[0].Cell(2, 1).GetValue().Value);
        }

        [Fact]
        public void ADestinationThatOpensTheSdkPackageWhileWritingSeesTheSavedChanges() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            document.Sheets[0].Cell(2, 1).SetValue("Changed");
            object? seen = null;
            var output = new CallbackStream {
                OnWrite = () => {
                    _ = document.OpenXmlDocument;
                    seen = document.Sheets[0].Cell(2, 1).GetValue().Value;
                }
            };
            document.Save(output);

            Assert.Equal("Changed", seen);
            using var second = new MemoryStream();
            document.Save(second);
            using var reread = ExcelDocument.Load(new MemoryStream(second.ToArray()));
            Assert.Equal("Changed", reread.Sheets[0].Cell(2, 1).GetValue().Value);
        }

        /// <summary>A seekable destination that can run code on its first write or refuse synchronous writes.</summary>
        private sealed class CallbackStream : Stream {
            internal MemoryStream Inner { get; } = new();
            internal Action? OnWrite { get; set; }
            internal bool RejectSynchronousWrites { get; set; }
            public override bool CanRead => true;
            public override bool CanSeek => true;
            public override bool CanWrite => true;
            public override long Length => Inner.Length;
            public override long Position { get => Inner.Position; set => Inner.Position = value; }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) => Inner.Read(buffer, offset, count);
            public override long Seek(long offset, SeekOrigin origin) => Inner.Seek(offset, origin);
            public override void SetLength(long value) => Inner.SetLength(value);
            public override void Write(byte[] buffer, int offset, int count) {
                if (RejectSynchronousWrites) throw new NotSupportedException("Synchronous writes are not supported.");
                Action? callback = OnWrite;
                OnWrite = null;
                callback?.Invoke();
                Inner.Write(buffer, offset, count);
            }
            public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
                Inner.Write(buffer, offset, count);
                return Task.CompletedTask;
            }
            public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        }

        [Fact]
        public void ASecondSaveStartsFromTheFirst() {
            byte[] source = CreateWorkbook();
            using var document = ExcelDocument.Load(new MemoryStream(source));
            document.Sheets[0].Cell(2, 2).SetValue("first");
            using (var first = new MemoryStream()) document.Save(first);
            document.Sheets[0].Cell(3, 2).SetValue("second");
            using var second = new MemoryStream();
            document.Save(second);
            Assert.Empty(document.OpenXmlSdkFallbacks);

            using var reread = ExcelDocument.Load(new MemoryStream(second.ToArray()));
            Assert.Equal("first", reread.Sheets[0].Cell(2, 2).GetValue().Value);
            Assert.Equal("second", reread.Sheets[0].Cell(3, 2).GetValue().Value);
        }

        [Fact]
        public void RepositoryWorkbooksSaveTheSameCellsAsTheSdk() {
            string folder = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents");
            var failures = new List<string>();
            int compared = 0;
            foreach (string path in Directory.EnumerateFiles(folder, "*.xlsx", SearchOption.AllDirectories).OrderBy(file => file, StringComparer.Ordinal)) {
                string? difference = CompareWithSdk(File.ReadAllBytes(path), out bool comparedFile);
                if (comparedFile) compared++;
                if (difference != null) failures.Add(path.Substring(folder.Length).TrimStart(Path.DirectorySeparatorChar) + ": " + difference);
            }
            Assert.True(compared > 100, "Only " + compared + " workbooks were compared.");
            Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures.Take(40)));
        }

        // Opt-in: OFFICEIMO_XLSX_CORPUS names a folder of workbooks from other producers.
        [Fact]
        public void ExternalCorpusWorkbooksSaveTheSameCellsAsTheSdk() {
            string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_XLSX_CORPUS");
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) return;
            var failures = new List<string>();
            int compared = 0;
            foreach (string path in Directory.EnumerateFiles(folder, "*.xlsx", SearchOption.AllDirectories).OrderBy(file => file, StringComparer.Ordinal)) {
                string? difference = CompareWithSdk(File.ReadAllBytes(path), out bool comparedFile);
                if (comparedFile) compared++;
                if (difference != null) failures.Add(path + ": " + difference);
            }
            Assert.True(failures.Count == 0, failures.Count + " differences in " + compared + " workbooks" + Environment.NewLine + string.Join(Environment.NewLine, failures.Take(60)));
        }

        /// <summary>Edits the first worksheet both ways and returns the first difference, or null.</summary>
        private static string? CompareWithSdk(byte[] source, out bool compared) {
            compared = false;
            try {
                using var probe = ExcelDocument.Load(new MemoryStream(source));
                if (!probe.IsSdkDeferred || probe.Sheets.Count == 0) return null;
            } catch (Exception) {
                return null;
            }

            byte[] expected;
            try {
                expected = SaveThroughSdk(source, ApplyCorpusEdits);
            } catch (Exception) {
                return null;
            }
            byte[] actual;
            bool owned;
            try {
                actual = SaveWithoutSdk(source, ApplyCorpusEdits, out owned);
            } catch (Exception exception) {
                return "owned path threw " + exception.GetType().Name + ": " + exception.Message;
            }
            // Only a workbook the owned path saved counts as compared.
            if (!owned) return null;
            compared = true;
            // The edited cells and the dimension match the SDK's save; the model's save also repairs damaged cells
            // elsewhere, which the owned save leaves as the source has them.
            var sourceCells = CellMapsOf(source);
            var expectedCells = CellMapsOf(expected);
            var actualCells = CellMapsOf(actual);
            if (expectedCells.Count != actualCells.Count) return "sheet count " + expectedCells.Count + " vs " + actualCells.Count;
            for (int index = 0; index < expectedCells.Count; index++) {
                // Source references are included, so a cell both saves drop is still compared with what was read.
                foreach (string reference in expectedCells[index].Keys.Union(actualCells[index].Keys).Union(sourceCells[index].Keys)) {
                    bool edited = index == 0 && (reference == "dimension" || CorpusEditReferences.Contains(reference));
                    var baseline = edited ? expectedCells[index] : sourceCells[index];
                    baseline.TryGetValue(reference, out string? want);
                    actualCells[index].TryGetValue(reference, out string? got);
                    if (!edited && reference == "dimension") continue;
                    if (edited && CorpusStyleOnlyReferences.Contains(reference) && got != null) {
                        sourceCells[index].TryGetValue(reference, out string? source2);
                        string Value(string? text) => text == null ? "n  " : text.Substring(0, text.IndexOf(StyleSeparator, StringComparison.Ordinal));
                        string Style(string? text) => text == null ? "" : text.Substring(text.IndexOf(StyleSeparator, StringComparison.Ordinal));
                        want = Value(source2) + Style(want);
                        got = Value(got) + Style(got);
                    }
                    if (want != got) return "sheet " + (index + 1) + " " + reference + ": " + (edited ? "SDK" : "source") + " [" + want + "] owned [" + got + "]";
                }
            }
            return null;
        }

        private static void ApplyEdits(ExcelDocument document) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 5).SetValue("Status");
            sheet.Cell(2, 5).SetValue("Reviewed");
            sheet.Cell(3, 5).SetValue("Reviewed");
            sheet.Cell(2, 2).SetValue(42.5);
            sheet.Cell(3, 3).SetValue(new DateTime(2025, 3, 4));
            sheet.Cell(4, 1).SetValue(true);
            sheet.Cell(10, 2).SetValue(12345678901L);
            sheet.Cell(10, 1).SetValue(7m);
            sheet.Cell(12, 3).SetValue((object)"boxed");
            sheet.Cell(1, 1).SetBold().SetFillColor(OfficeColor.ParseHex("DDEBF7"));
            sheet.Cell(1, 2).SetBold();
            sheet.Cell(2, 2).SetNumberFormat("0.000");
            sheet.Cell(3, 1).SetItalic().SetFontSize(OfficeLength.Points(14)).SetColor(OfficeColor.ParseHex("C00000"));
            sheet.Cell(4, 1).SetBorder(ExcelBorderStyle.Thin).SetUnderline(ExcelUnderlineStyle.Double);
            sheet.Cell(4, 2).SetBorder(ExcelCellBorders.Bottom, ExcelBorderStyle.Medium, OfficeColor.ParseHex("1F4E79"));
            sheet.Cell(5, 5).SetHorizontalAlignment(ExcelHorizontalAlignment.Center).SetWrapText();
            sheet.Cell(3, 3).SetVerticalAlignment(ExcelVerticalAlignment.Top).SetStrikethrough().SetFontFamily("Arial");
            sheet.Cell(2, 1).SetTextRotation(45).SetShrinkToFit().SetSuperscript();
            sheet.Cell(13, 1).SetNumberFormat("#,##0.00").SetValue(1234.5);
            sheet.Cell(14, 1).SetValue(new DateTime(2025, 1, 2)).SetBold();
        }

        private static readonly HashSet<string> CorpusEditReferences = new(StringComparer.Ordinal) { "A1", "B2", "C3", "D4", "A5", "C2000", "B6", "C7" };
        // Cells the corpus edits only restyle: their value is the source's, which the model's save may also repair.
        private static readonly HashSet<string> CorpusStyleOnlyReferences = new(StringComparer.Ordinal) { "B6", "C7" };
        private const string StyleSeparator = " | ";

        private static void ApplyCorpusEdits(ExcelDocument document) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1).SetValue("Edited A1");
            sheet.Cell(2, 2).SetValue(3.25);
            sheet.Cell(3, 3).SetValue(new DateTime(2024, 2, 29));
            sheet.Cell(4, 4).SetValue(false);
            sheet.Cell(5, 1).SetValue("Edited A1");
            sheet.Cell(2000, 3).SetValue("New row");
            sheet.Cell(1, 1).SetBold();
            sheet.Cell(2, 2).SetNumberFormat("0.0%");
            sheet.Cell(6, 2).SetFillColor(OfficeColor.ParseHex("FFF2CC")).SetBorder(ExcelBorderStyle.Thin);
            sheet.Cell(7, 3).SetHorizontalAlignment(ExcelHorizontalAlignment.Right);
        }

        private static byte[] SaveWithoutSdk(byte[] source, Action<ExcelDocument> edit, out bool owned) {
            using var document = ExcelDocument.Load(new MemoryStream(source));
            edit(document);
            using var output = new MemoryStream();
            document.Save(output);
            owned = document.OpenXmlSdkFallbacks.Count == 0;
            return output.ToArray();
        }

        private static byte[] SaveThroughSdk(byte[] source, Action<ExcelDocument> edit) {
            using var document = ExcelDocument.Load(new MemoryStream(source));
            document.OpenDeferredSdk("test");
            edit(document);
            using var output = new MemoryStream();
            document.Save(output);
            return output.ToArray();
        }

        private static byte[] EntryBytes(byte[] package, string name) {
            using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            return ContentOf(archive.GetEntry(name) ?? throw new InvalidOperationException(name + " is missing."));
        }

        private static byte[] WithEntry(byte[] package, string name, string xml) {
            using var stream = new MemoryStream();
            stream.Write(package, 0, package.Length);
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry? entry = archive.GetEntry(name);
                Assert.NotNull(entry);
                entry!.Delete();
                using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(xml);
            }
            return stream.ToArray();
        }

        private static byte[] CreateWorkbook() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet data = document.AddSheet("Data");
                data.Cell(1, 1).SetValue("Name");
                data.Cell(1, 2).SetValue("Amount");
                data.Cell(1, 3).SetValue("Date");
                for (int row = 2; row <= 4; row++) {
                    data.Cell(row, 1).SetValue("Item " + row);
                    data.Cell(row, 2).SetValue(row * 1.5);
                    data.Cell(row, 3).SetValue(new DateTime(2024, 1, row));
                }
                data.Cell(2, 2).SetNumberFormat("#,##0.00");
                ExcelSheet second = document.AddSheet("Second");
                second.Cell(1, 1).SetValue("Untouched");
                document.Save();
            }
            return stream.ToArray();
        }

        private static byte[] CreateEmptySheetWorkbook() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.AddSheet("Data");
                document.Save();
            }
            return stream.ToArray();
        }

        private static List<string> ChangedEntries(byte[] before, byte[] after) {
            using var left = new ZipArchive(new MemoryStream(before), ZipArchiveMode.Read);
            using var right = new ZipArchive(new MemoryStream(after), ZipArchiveMode.Read);
            var changed = new List<string>();
            foreach (ZipArchiveEntry entry in right.Entries) {
                ZipArchiveEntry? original = left.GetEntry(entry.FullName);
                if (original == null || !ContentOf(original).SequenceEqual(ContentOf(entry))) {
                    changed.Add(entry.FullName);
                    continue;
                }
#if NET8_0_OR_GREATER
                // Only the .NET 8+ writer copies untouched members' compressed bytes; older targets re-deflate them.
                if (original.CompressedLength != entry.CompressedLength) changed.Add(entry.FullName);
#endif
            }
            // A part the save dropped counts as changed too.
            foreach (ZipArchiveEntry entry in left.Entries) {
                if (right.GetEntry(entry.FullName) == null) changed.Add(entry.FullName);
            }
            return changed;
        }

        private static byte[] ContentOf(ZipArchiveEntry entry) {
            using Stream stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }

        private static string Flag(string value) => value == "true" ? "1" : value == "false" ? "0" : value;

        /// <summary>An element's content with attributes and children in a fixed order, so equal formats compare equal.</summary>
        private static string Canonical(XElement element) =>
            "<" + element.Name.LocalName + string.Concat(element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                .Select(attribute => " " + attribute.Name.LocalName + "=" + Flag(attribute.Value)).OrderBy(text => text, StringComparer.Ordinal))
            + ">" + string.Concat(element.Elements().Select(Canonical).OrderBy(text => text, StringComparer.Ordinal))
            + (element.HasElements ? string.Empty : element.Value) + "</>";

        private static string DescribeFormula(XElement? formula) => formula == null
            ? string.Empty
            : "f(" + string.Join(",", formula.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                .Select(attribute => attribute.Name.LocalName + "=" + attribute.Value).OrderBy(text => text, StringComparer.Ordinal)) + ")" + formula.Value;

        /// <summary>
        /// Each worksheet's dimension and cells, with shared strings resolved and styles described by their format,
        /// so the comparison does not depend on where in the tables a string or format sits.
        /// </summary>
        private static List<string> CellsOf(byte[] package) =>
            CellMapsOf(package).Select(map => string.Join("\n", map.Select(pair => pair.Key + " " + pair.Value))).ToList();

        private static List<SortedDictionary<string, string>> CellMapsOf(byte[] package) {
            using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            XDocument Load(string name) {
                using Stream stream = archive.GetEntry(name)!.Open();
                return XDocument.Load(stream);
            }
            static string Resolve(string directory, string target) {
                var segments = new List<string>();
                string path = target.StartsWith("/", StringComparison.Ordinal) ? target.Substring(1) : (directory.Length == 0 ? target : directory + "/" + target);
                foreach (string segment in path.Split('/')) {
                    if (segment == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); }
                    else if (segment.Length > 0 && segment != ".") segments.Add(Uri.UnescapeDataString(segment));
                }
                return string.Join("/", segments);
            }

            XDocument packageRelationships = Load("_rels/.rels");
            string workbookName = Resolve("", packageRelationships.Root!.Elements(PackageRelationships + "Relationship")
                .First(element => ((string)element.Attribute("Type")!).EndsWith("/officeDocument", StringComparison.Ordinal)).Attribute("Target")!.Value);
            string workbookDirectory = workbookName.Contains('/') ? workbookName.Substring(0, workbookName.LastIndexOf('/')) : "";
            string relationshipsName = (workbookDirectory.Length == 0 ? "" : workbookDirectory + "/") + "_rels/" + Path.GetFileName(workbookName) + ".rels";
            var relationships = Load(relationshipsName).Root!.Elements(PackageRelationships + "Relationship").ToList();
            string? Target(string type) => relationships.FirstOrDefault(element => ((string)element.Attribute("Type")!).EndsWith("/" + type, StringComparison.Ordinal))
                ?.Attribute("Target")?.Value is { } target ? Resolve(workbookDirectory, target) : null;

            var strings = new List<string>();
            if (Target("sharedStrings") is { } sharedStrings && archive.GetEntry(sharedStrings) != null) {
                foreach (XElement item in Load(sharedStrings).Root!.Elements(Main + "si")) {
                    strings.Add(item.Element(Main + "t")?.Value ?? string.Concat(item.Elements(Main + "r").Select(run => run.Element(Main + "t")?.Value)));
                }
            }
            var formats = new List<string>();
            var codes = new Dictionary<string, string>();
            if (Target("styles") is { } styles && archive.GetEntry(styles) != null) {
                XElement root = Load(styles).Root!;
                foreach (XElement format in root.Element(Main + "numFmts")?.Elements(Main + "numFmt") ?? Enumerable.Empty<XElement>()) {
                    codes[(string)format.Attribute("numFmtId")!] = (string?)format.Attribute("formatCode") ?? "";
                }
                List<string> Entries(string collection, string item) =>
                    (root.Element(Main + collection)?.Elements(Main + item) ?? Enumerable.Empty<XElement>()).Select(Canonical).ToList();
                var fonts = Entries("fonts", "font");
                var fills = Entries("fills", "fill");
                var borders = Entries("borders", "border");
                string Entry(List<string> entries, string? id) =>
                    int.TryParse(id ?? "0", NumberStyles.None, CultureInfo.InvariantCulture, out int index) && index < entries.Count ? entries[index] : "#" + id;
                int cellStyleFormats = root.Element(Main + "cellStyleXfs")?.Elements(Main + "xf").Count() ?? 0;
                foreach (XElement format in root.Element(Main + "cellXfs")?.Elements(Main + "xf") ?? Enumerable.Empty<XElement>()) {
                    string id = (string?)format.Attribute("numFmtId") ?? "0";
                    string[] references = { "numFmtId", "fontId", "fillId", "borderId", "xfId" };
                    // A cell style past the end of the table is the Normal style; the model's save also repairs it to 0.
                    string cellStyle = int.TryParse((string?)format.Attribute("xfId") ?? "0", NumberStyles.None, CultureInfo.InvariantCulture, out int xfId)
                        && xfId < cellStyleFormats ? xfId.ToString(CultureInfo.InvariantCulture) : "0";
                    formats.Add("xfId=" + cellStyle + "," + string.Join(",", format.Attributes().Where(attribute => !references.Contains(attribute.Name.LocalName))
                        .Select(attribute => attribute.Name.LocalName + "=" + Flag(attribute.Value))
                        .OrderBy(text => text, StringComparer.Ordinal))
                        + ",fmt=" + (codes.TryGetValue(id, out string? code) ? code : id)
                        + ",font=" + Entry(fonts, (string?)format.Attribute("fontId"))
                        + ",fill=" + Entry(fills, (string?)format.Attribute("fillId"))
                        + ",border=" + Entry(borders, (string?)format.Attribute("borderId"))
                        + "," + string.Concat(format.Elements().Select(Canonical)));
                }
            }

            var result = new List<SortedDictionary<string, string>>();
            XDocument workbook = Load(workbookName);
            foreach (XElement sheet in workbook.Root!.Element(Main + "sheets")!.Elements(Main + "sheet")) {
                string id = (string)sheet.Attribute(Relationships + "id")!;
                XElement? relationship = relationships.FirstOrDefault(element => (string?)element.Attribute("Id") == id);
                if (relationship == null || !((string)relationship.Attribute("Type")!).EndsWith("/worksheet", StringComparison.Ordinal)) continue;
                XDocument worksheet = Load(Resolve(workbookDirectory, (string)relationship.Attribute("Target")!));
                var map = new SortedDictionary<string, string>(StringComparer.Ordinal) {
                    ["dimension"] = (string?)worksheet.Root!.Element(Main + "dimension")?.Attribute("ref") ?? ""
                };
                foreach (XElement row in worksheet.Root.Element(Main + "sheetData")?.Elements(Main + "row") ?? Enumerable.Empty<XElement>()) {
                    foreach (XElement cell in row.Elements(Main + "c")) {
                        string type = (string?)cell.Attribute("t") ?? "n";
                        string? value = cell.Element(Main + "v")?.Value;
                        if (type == "s" && int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int index) && index < strings.Count) value = "'" + strings[index];
                        if (type == "inlineStr" && cell.Element(Main + "is") is { } inline) {
                            value = "'" + (inline.Element(Main + "t")?.Value ?? string.Concat(inline.Elements(Main + "r").Select(run => run.Element(Main + "t")?.Value)));
                        }
                        // A cell without a style index takes format 0.
                        string styleText = (string?)cell.Attribute("s") ?? "0";
                        string style = int.TryParse(styleText, out int styleIndex) && styleIndex < formats.Count ? formats[styleIndex] : styleText;
                        map[(string?)cell.Attribute("r") ?? "?"] = type + " " + value + " " + DescribeFormula(cell.Element(Main + "f")) + StyleSeparator + style;
                    }
                }
                result.Add(map);
            }
            return result;
        }
    }
}
