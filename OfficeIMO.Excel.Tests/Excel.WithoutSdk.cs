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
            Assert.Equal(new[] { "docProps/core.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml" }, changed.OrderBy(name => name, StringComparer.Ordinal));
        }

        [Fact]
        public void AnOperationTheOwnedPathDoesNotCoverOpensTheSdkWithTheValuesWritten() {
            byte[] source = CreateWorkbook();
            byte[] expected = SaveThroughSdk(source, document => {
                ApplyEdits(document);
                document.Sheets[0].Cell(1, 1).SetBold();
            });
            byte[] actual;
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                ApplyEdits(document);
                document.Sheets[0].Cell(1, 1).SetBold();
                Assert.Single(document.OpenXmlSdkFallbacks);
                using var output = new MemoryStream();
                document.Save(output);
                actual = output.ToArray();
            }
            Assert.Equal(CellsOf(expected), CellsOf(actual));
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
                if (difference != null) failures.Add(Path.GetRelativePath(folder, path) + ": " + difference);
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
            compared = true;
            byte[] actual;
            bool owned;
            try {
                actual = SaveWithoutSdk(source, ApplyCorpusEdits, out owned);
            } catch (Exception exception) {
                return "owned path threw " + exception.GetType().Name + ": " + exception.Message;
            }
            if (!owned) return null;
            // The edited cells and the dimension match the SDK's save; the model's save also repairs damaged cells
            // elsewhere, which the owned save leaves as the source has them.
            var sourceCells = CellMapsOf(source);
            var expectedCells = CellMapsOf(expected);
            var actualCells = CellMapsOf(actual);
            if (expectedCells.Count != actualCells.Count) return "sheet count " + expectedCells.Count + " vs " + actualCells.Count;
            for (int index = 0; index < expectedCells.Count; index++) {
                foreach (string reference in expectedCells[index].Keys.Union(actualCells[index].Keys)) {
                    bool edited = index == 0 && (reference == "dimension" || CorpusEditReferences.Contains(reference));
                    var baseline = edited ? expectedCells[index] : sourceCells[index];
                    baseline.TryGetValue(reference, out string? want);
                    actualCells[index].TryGetValue(reference, out string? got);
                    if (!edited && reference == "dimension") continue;
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
        }

        private static readonly HashSet<string> CorpusEditReferences = new(StringComparer.Ordinal) { "A1", "B2", "C3", "D4", "A5", "C2000" };

        private static void ApplyCorpusEdits(ExcelDocument document) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1).SetValue("Edited A1");
            sheet.Cell(2, 2).SetValue(3.25);
            sheet.Cell(3, 3).SetValue(new DateTime(2024, 2, 29));
            sheet.Cell(4, 4).SetValue(false);
            sheet.Cell(5, 1).SetValue("Edited A1");
            sheet.Cell(2000, 3).SetValue("New row");
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

        private static List<string> ChangedEntries(byte[] before, byte[] after) {
            using var left = new ZipArchive(new MemoryStream(before), ZipArchiveMode.Read);
            using var right = new ZipArchive(new MemoryStream(after), ZipArchiveMode.Read);
            var changed = new List<string>();
            foreach (ZipArchiveEntry entry in right.Entries) {
                ZipArchiveEntry? original = left.GetEntry(entry.FullName);
                if (original == null || original.Crc32 != entry.Crc32 || original.CompressedLength != entry.CompressedLength) changed.Add(entry.FullName);
            }
            return changed;
        }

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
                foreach (XElement format in root.Element(Main + "cellXfs")?.Elements(Main + "xf") ?? Enumerable.Empty<XElement>()) {
                    string id = (string?)format.Attribute("numFmtId") ?? "0";
                    formats.Add(string.Join(",", format.Attributes().Where(attribute => attribute.Name.LocalName != "numFmtId")
                        .Select(attribute => attribute.Name.LocalName + "=" + (attribute.Value == "true" ? "1" : attribute.Value == "false" ? "0" : attribute.Value))
                        .OrderBy(text => text, StringComparer.Ordinal)) + ",fmt=" + (codes.TryGetValue(id, out string? code) ? code : id) + "," + format.Elements().Count());
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
                        string style = (string?)cell.Attribute("s") is { } styleText && int.TryParse(styleText, out int styleIndex) && styleIndex < formats.Count
                            ? formats[styleIndex] : (string?)cell.Attribute("s") ?? "";
                        map[(string?)cell.Attribute("r") ?? "?"] = type + " " + value + " " + DescribeFormula(cell.Element(Main + "f")) + " " + style;
                    }
                }
                result.Add(map);
            }
            return result;
        }
    }
}
