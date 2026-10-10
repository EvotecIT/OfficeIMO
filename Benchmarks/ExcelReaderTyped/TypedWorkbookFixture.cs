using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>The materialized model used by every typed benchmark.</summary>
    public sealed class TypedRecord {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    internal static class TypedWorkbookFixture {
        internal static readonly string[] Headers = ["Name", "Id", "Date", "Value"];
        private static readonly string[] Names = ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];

        internal static TypedRecord ExpectedRecord(int index) => new() {
            Name = Names[index % Names.Length],
            Id = index,
            Date = DateTime.FromOADate(45292 + index % 3650 + 0.25),
            Value = index * 1.5,
        };

        internal static long Accumulate(TypedRecord record) =>
            record.Id + (long)record.Value + (record.Name?.Length ?? 0) + record.Date.Ticks;

        internal static long ExpectedChecksum(int rowCount) {
            long result = 0;
            for (int row = 1; row <= rowCount; row++) result = unchecked(result + Accumulate(ExpectedRecord(row)));
            return result;
        }

        internal static void ValidateRecord(TypedRecord record, int index) {
            TypedRecord expected = ExpectedRecord(index);
            if (record.Name != expected.Name || record.Id != expected.Id || record.Date != expected.Date || record.Value != expected.Value)
                throw new InvalidDataException($"Typed values differ at data row {index}.");
        }

        // Same record values, writer API, sheet name, and default writer options as
        // the upstream BuildTypedAsync workload identified in README.md.
        internal static async Task<byte[]> CreateAsync(int rowCount, string shape) {
            List<TypedRecord> records = new List<TypedRecord>(rowCount);
            for (int row = 1; row <= rowCount; row++) records.Add(ExpectedRecord(row));
            await using MemoryStream stream = new MemoryStream();
            await using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true,
                options: shape == "SharedStrings" ? new XlsxWriterOptions { UseSharedStrings = true } : null)) {
                await using XlsxSheetWriter sheet = workbook.AddSheet("S1");
                await sheet.WriteRecordsAsync(records, ExcelRecordLayout.FromAttributes<TypedRecord>());
            }
            byte[] original = stream.ToArray();
            byte[] bytes = shape is "Original" or "SharedStrings" ? original : Rewrite(original, rowCount, shape);
            BenchmarkInput.WriteWorkbookFixtureIdentity($"read/typed/Xlsx/{shape}/dataRows={rowCount}", bytes);
            return bytes;
        }

        internal static async Task<byte[]> CreateXlsbAsync(int rowCount) {
            List<TypedRecord> records = new List<TypedRecord>(rowCount);
            for (int row = 1; row <= rowCount; row++) records.Add(ExpectedRecord(row));
            await using MemoryStream stream = new MemoryStream();
            await using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(stream, leaveOpen: true)) {
                await using XlsbSheetWriter sheet = workbook.AddSheet("S1");
                await sheet.WriteRecordsAsync(records, ExcelRecordLayout.FromAttributes<TypedRecord>());
            }
            byte[] bytes = stream.ToArray();
            BenchmarkInput.WriteWorkbookFixtureIdentity($"read/typed/Xlsb/dataRows={rowCount}", bytes);
            return bytes;
        }

        internal static string ValidateShape(byte[] workbook, int rowCount, string shape) {
            using MemoryStream stream = new MemoryStream(workbook, writable: false);
            using ZipArchive package = new ZipArchive(stream, ZipArchiveMode.Read);
            if (package.Entries.Count(entry => entry.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal)
                    && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)) != 1)
                throw new InvalidDataException("The generated workbook must have exactly one worksheet.");
            ZipArchiveEntry part = package.GetEntry("xl/worksheets/sheet1.xml")
                ?? throw new InvalidDataException("The generated worksheet is missing.");
            using XmlReader reader = XmlReader.Create(part.Open(), new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = true,
            });
            int rows = 0, cells = 0, rowReferences = 0, cellReferences = 0, sharedStrings = 0, inlineStrings = 0;
            string? dimension = null;
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "row") {
                    rows++;
                    string? rowReference = reader.GetAttribute("r");
                    if (rowReference != null) {
                        rowReferences++;
                        if (rowReference != rows.ToString(CultureInfo.InvariantCulture))
                            throw new InvalidDataException("The row-reference variant is not sequential.");
                    }
                } else if (reader.LocalName == "c") {
                    cells++;
                    if (reader.GetAttribute("r") != null) cellReferences++;
                    if (reader.GetAttribute("t") == "s") sharedStrings++;
                    if (reader.GetAttribute("t") == "inlineStr") inlineStrings++;
                } else if (reader.LocalName == "dimension") {
                    if (dimension != null) throw new InvalidDataException("Duplicate worksheet dimension.");
                    dimension = reader.GetAttribute("ref");
                }
            }
            int expectedRows = rowCount + 1;
            bool expectedDimension = shape == "Dimension";
            if (rows != expectedRows || cells != expectedRows * Headers.Length || cellReferences != 0
                || rowReferences != (shape == "RowReferences" ? expectedRows : 0)
                || (expectedDimension ? dimension != $"A1:D{expectedRows}" : dimension != null))
                throw new InvalidDataException("The generated worksheet does not match the selected shape.");
            bool usesSharedStrings = shape == "SharedStrings";
            int textCells = rowCount + Headers.Length;
            if (sharedStrings != (usesSharedStrings ? textCells : 0) || inlineStrings != (usesSharedStrings ? 0 : textCells))
                throw new InvalidDataException("The string storage does not match the selected shape.");
            ZipArchiveEntry? sharedPart = package.GetEntry("xl/sharedStrings.xml");
            if (usesSharedStrings) {
                if (sharedPart == null) throw new InvalidDataException("The shared string table is missing.");
                using XmlReader sharedReader = XmlReader.Create(sharedPart.Open(), new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, CloseInput = true,
                });
                string[] values = XDocument.Load(sharedReader).Descendants().Where(element => element.Name.LocalName == "si")
                    .Select(element => element.Value).ToArray();
                IEnumerable<string> expected = Headers.Concat(Enumerable.Range(1, Math.Min(rowCount, Names.Length)).Select(index => Names[index % Names.Length]));
                if (values.Length != Math.Min(rowCount, Names.Length) + Headers.Length || !values.ToHashSet(StringComparer.Ordinal).SetEquals(expected))
                    throw new InvalidDataException("The shared string table values differ.");
            } else if (sharedPart != null) {
                throw new InvalidDataException("The inline input unexpectedly contains a shared string table.");
            }
            return $"rows={rows}, cells={cells}, rowReferences={rowReferences}, cellReferences={cellReferences}, "
                + $"dimension={dimension ?? "none"}, sharedStringCells={sharedStrings}, inlineStringCells={inlineStrings}, "
                + $"packageBytes={workbook.Length}, worksheetBytes={part.Length}";
        }

        private static byte[] Rewrite(byte[] original, int rowCount, string shape) {
            using MemoryStream sourceStream = new MemoryStream(original, writable: false);
            using ZipArchive source = new ZipArchive(sourceStream, ZipArchiveMode.Read);
            using MemoryStream destinationStream = new MemoryStream();
            using (ZipArchive destination = new ZipArchive(destinationStream, ZipArchiveMode.Create, leaveOpen: true)) {
                foreach (ZipArchiveEntry entry in source.Entries) {
                    using Stream input = entry.Open();
                    using Stream output = destination.CreateEntry(entry.FullName, CompressionLevel.Fastest).Open();
                    if (entry.FullName != "xl/worksheets/sheet1.xml") {
                        input.CopyTo(output);
                        continue;
                    }
                    using StreamReader reader = new StreamReader(input);
                    string xml = reader.ReadToEnd();
                    if (shape == "RowReferences") {
                        int row = 0;
                        xml = Regex.Replace(xml, @"<row(?=[ >])", _ => $"<row r=\"{++row}\"");
                    } else if (shape == "Dimension") {
                        xml = xml.Replace("<sheetData>", $"<dimension ref=\"A1:D{rowCount + 1}\"/><sheetData>", StringComparison.Ordinal);
                    } else {
                        throw new ArgumentException($"Unknown worksheet shape '{shape}'.", nameof(shape));
                    }
                    using StreamWriter writer = new StreamWriter(output, new UTF8Encoding(false));
                    writer.Write(xml);
                }
            }
            return destinationStream.ToArray();
        }
    }
}
