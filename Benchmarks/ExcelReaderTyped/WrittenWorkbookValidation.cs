using System.Globalization;
using System.IO.Compression;
using System.IO.Hashing;
using System.Xml;
using System.Xml.Linq;
using ExcelReader.Core.Parser;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

internal static class WrittenWorkbookValidation {
    private const string Worksheet = "xl/worksheets/sheet1.xml";

    internal static void Validate(byte[] bytes, int rowCount, bool officeIMO, bool sharedStrings = false, bool? includeReferences = null) {
        string engine = officeIMO ? "OfficeIMO" : "ExcelReader";
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = new ZipArchive(stream, ZipArchiveMode.Read);
        ValidateEntries(package);
        ValidateRelationships(package);
        string[] required = ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels", "xl/styles.xml", Worksheet];
        foreach (string part in required) {
            if (package.GetEntry(part) == null) throw new InvalidDataException($"{engine} output is missing {part}.");
        }
        if (officeIMO && (package.GetEntry("docProps/core.xml") == null || package.GetEntry("docProps/app.xml") == null))
            throw new InvalidDataException("OfficeIMO output is missing its ordinary document properties.");
        XDocument workbook = ReadXml(package.GetEntry("xl/workbook.xml")!);
        XElement[] sheets = workbook.Descendants().Where(element => element.Name.LocalName == "sheet").ToArray();
        if (sheets.Length != 1 || (string?)sheets[0].Attribute("name") != (officeIMO ? "Data" : "S1"))
            throw new InvalidDataException("The written sheet metadata differs.");
        string? date1904 = (string?)workbook.Descendants().SingleOrDefault(element => element.Name.LocalName == "workbookPr")?.Attribute("date1904");
        if (date1904 is "1" or "true") throw new InvalidDataException("The writer changed the date system.");
        XDocument styles = ReadXml(package.GetEntry("xl/styles.xml")!);
        XElement[] cellFormats = styles.Descendants().Single(element => element.Name.LocalName == "cellXfs").Elements().ToArray();
        ZipArchiveEntry? sharedPart = package.GetEntry("xl/sharedStrings.xml");
        string[] sharedValues = sharedPart == null ? [] : ReadXml(sharedPart).Descendants()
            .Where(element => element.Name.LocalName == "si").Select(element => element.Value).ToArray();
        if (sharedStrings != (sharedValues.Length != 0)) throw new InvalidDataException("The writer did not honor the selected shared-string policy.");
        string description = ValidateWorksheet(package.GetEntry(Worksheet)!, rowCount, includeReferences ?? officeIMO,
            officeIMO, cellFormats, sharedStrings, sharedValues);
        ValidateIndependentReaders(bytes, rowCount);
        BenchmarkInput.WriteWorkbookFixtureIdentity($"validated-output/{engine}/Xlsx/dataRows={rowCount}/"
            + $"sharedStrings={sharedStrings}/references={includeReferences ?? officeIMO}", bytes);
        Console.WriteLine($"Validated {engine} write: {description}, packageBytes={bytes.Length}, "
            + $"checksum={TypedWorkbookFixture.ExpectedChecksum(rowCount)}, independentReaders=ExcelReader+Sylvan.");
    }

    internal static void ValidateEntries(ZipArchive package) {
        var names = new HashSet<string>(StringComparer.Ordinal);
        byte[] buffer = new byte[8192];
        foreach (ZipArchiveEntry entry in package.Entries) {
            if (!names.Add(entry.FullName)) throw new InvalidDataException("Duplicate ZIP entry.");
            using (Stream input = entry.Open()) {
                var crc = new Crc32();
                long read = 0;
                int count;
                while ((count = input.Read(buffer, 0, buffer.Length)) != 0) {
                    read += count;
                    crc.Append(buffer.AsSpan(0, count));
                }
                if (read != entry.Length || crc.GetCurrentHashAsUInt32() != entry.Crc32)
                    throw new InvalidDataException("A ZIP entry's payload length or CRC differs.");
            }
            if (entry.FullName.EndsWith(".xml", StringComparison.Ordinal) || entry.FullName.EndsWith(".rels", StringComparison.Ordinal)) {
                using var reader = CreateReader(entry.Open());
                while (reader.Read()) { }
            }
        }
    }

    internal static void ValidateRelationships(ZipArchive package) {
        foreach (ZipArchiveEntry entry in package.Entries.Where(item => item.FullName.EndsWith(".rels", StringComparison.Ordinal))) {
            string source = entry.FullName == "_rels/.rels" ? "" : entry.FullName.Replace("/_rels/", "/", StringComparison.Ordinal)[..^5];
            var sourceUri = new Uri("https://package.invalid/" + source);
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (XElement relationship in ReadXml(entry).Root!.Elements()) {
                string? id = (string?)relationship.Attribute("Id"), target = (string?)relationship.Attribute("Target");
                if (id == null || target == null || !ids.Add(id) || relationship.Attribute("Type") == null
                    || (string?)relationship.Attribute("TargetMode") == "External")
                    throw new InvalidDataException("Invalid or unexpected external package relationship.");
                var resolved = new Uri(sourceUri, target);
                string part = Uri.UnescapeDataString(resolved.AbsolutePath.TrimStart('/'));
                if (resolved.Host != sourceUri.Host || package.GetEntry(part) == null)
                    throw new InvalidDataException("A package relationship targets a missing part.");
            }
        }
        foreach (XElement part in ReadXml(package.GetEntry("[Content_Types].xml")!).Root!.Elements().Where(item => item.Name.LocalName == "Override")) {
            string? name = (string?)part.Attribute("PartName");
            if (name == null || part.Attribute("ContentType") == null || package.GetEntry(name.TrimStart('/')) == null)
                throw new InvalidDataException("An override content type targets a missing part.");
        }
    }

    private static string ValidateWorksheet(ZipArchiveEntry part, int rowCount, bool includeReferences,
        bool headerReferences, XElement[] styles, bool sharedStrings, string[] sharedValues) {
        using var reader = CreateReader(part.Open());
        int rows = 0, column = 0, rowReferences = 0, cellReferences = 0, dateCells = 0;
        string? dimension = null;
        while (reader.Read()) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row") {
                if (column != 4) throw new InvalidDataException("Incorrect written row width.");
                continue;
            }
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.LocalName == "dimension") {
                dimension = reader.GetAttribute("ref");
            } else if (reader.LocalName == "row") {
                rows++;
                column = 0;
                string? rowReference = reader.GetAttribute("r");
                if (rowReference != null) {
                    rowReferences++;
                    if (rowReference != rows.ToString(CultureInfo.InvariantCulture)) throw new InvalidDataException("Written row order differs.");
                }
            } else if (reader.LocalName == "c") {
                int currentColumn = column++;
                if (currentColumn >= 4 || rows > rowCount + 1) throw new InvalidDataException("Extra written cells.");
                string? reference = reader.GetAttribute("r");
                string? type = reader.GetAttribute("t");
                string? style = reader.GetAttribute("s");
                if (reference != null) {
                    cellReferences++;
                    if (reference != $"{(char)('A' + currentColumn)}{rows}") throw new InvalidDataException("Incorrect cell reference.");
                }
                string? value = null;
                using (var cell = reader.ReadSubtree()) {
                    while (cell.Read()) {
                        if (cell.NodeType == XmlNodeType.Element && cell.LocalName is "v" or "t")
                            value = cell.ReadElementContentAsString();
                    }
                }
                if (type == "s") {
                    if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int sharedIndex)
                        || sharedIndex < 0 || sharedIndex >= sharedValues.Length)
                        throw new InvalidDataException("Invalid shared-string reference.");
                    value = sharedValues[sharedIndex];
                }
                if (rows == 1) {
                    if (type != (sharedStrings ? "s" : "inlineStr") || value != TypedWorkbookFixture.Headers[currentColumn])
                        throw new InvalidDataException("Incorrect written header.");
                    continue;
                }
                TypedRecord expected = TypedWorkbookFixture.ExpectedRecord(rows - 1);
                if (currentColumn == 0) {
                    if (type != (sharedStrings ? "s" : "inlineStr") || value != expected.Name) throw new InvalidDataException("Incorrect written name.");
                } else {
                    if (type is not (null or "n") || !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
                        throw new InvalidDataException("Incorrect numeric cell serialization.");
                    double expectedNumber = currentColumn switch { 1 => expected.Id, 2 => expected.Date.ToOADate(), _ => expected.Value };
                    if (number != expectedNumber) throw new InvalidDataException("Incorrect written numeric/date value.");
                    if (currentColumn == 2) {
                        if (!int.TryParse(style, out int styleIndex) || styleIndex < 0 || styleIndex >= styles.Length
                            || !uint.TryParse((string?)styles[styleIndex].Attribute("numFmtId"), out uint numberFormat) || numberFormat == 0)
                            throw new InvalidDataException("The written date has no valid number format.");
                        dateCells++;
                    }
                }
            }
        }
        int expectedRows = rowCount + 1;
        int expectedRowReferences = includeReferences ? expectedRows : headerReferences ? 1 : 0;
        int expectedCellReferences = includeReferences ? expectedRows * 4 : headerReferences ? 4 : 0;
        if (rows != expectedRows || dateCells != rowCount || rowReferences != expectedRowReferences
            || cellReferences != expectedCellReferences)
            throw new InvalidDataException("Written rows, date styles, or coordinate policy differs.");
        if (dimension != null && dimension != $"A1:D{expectedRows}") throw new InvalidDataException("Incorrect written dimension.");
        return $"rows={rows}, dateCells={dateCells}, rowReferences={rowReferences}, cellReferences={cellReferences}, "
            + $"dimension={dimension ?? "none"}, worksheetBytes={part.Length}";
    }

    private static void ValidateIndependentReaders(byte[] bytes, int rowCount) {
        using (var stream = new MemoryStream(bytes, writable: false)) {
            using var workbook = ExcelReaderApi.FromXlsx(stream);
            if (workbook.SheetCount != 1) throw new InvalidDataException("Independent reader returned an incorrect sheet count.");
            int count = 0;
            foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet))
                TypedWorkbookFixture.ValidateRecord(record, ++count);
            if (count != rowCount) throw new InvalidDataException("Independent reader returned an incorrect row count.");
        }
        using var input = new MemoryStream(bytes, writable: false);
        using var reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(input, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
        if (reader.FieldCount != 4) throw new InvalidDataException("Independent reader returned an incorrect width.");
        for (int column = 0; column < 4; column++) {
            if (reader.GetName(column) != TypedWorkbookFixture.Headers[column]) throw new InvalidDataException("Independent reader returned an incorrect header.");
        }
        int rows = 0;
        while (reader.Read()) {
            if (reader.GetFormat(2)?.Kind != FormatKind.Date) throw new InvalidDataException("Independent reader did not recognize the date style.");
            TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = reader.GetString(0), Id = reader.GetInt32(1),
                Date = reader.GetDateTime(2), Value = reader.GetDouble(3) }, ++rows);
        }
        if (rows != rowCount || reader.NextResult()) throw new InvalidDataException("Independent reader returned an incorrect row/sheet count.");
    }

    private static XDocument ReadXml(ZipArchiveEntry entry) {
        using var reader = CreateReader(entry.Open());
        return XDocument.Load(reader);
    }

    private static XmlReader CreateReader(Stream stream) => XmlReader.Create(stream, new XmlReaderSettings {
        DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, CloseInput = true,
    });
}
