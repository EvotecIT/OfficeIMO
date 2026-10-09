using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Creates fixed-identity worksheet shapes without invoking a measured writer.</summary>
internal static class ExcelWorksheetPreparationFixture {
    private const string WorksheetPartName = "xl/worksheets/sheet1.xml";
    private const string XmlDeclaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly UTF8Encoding Utf8 = new(false, true);
    private static readonly DateTimeOffset FixedTimestamp = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
    private static readonly string[] PartNames = [
        "[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", WorksheetPartName
    ];

    internal static double Value(int id) => 1000D + id * 0.25D;

    internal static byte[] Create(int dataRows, WorksheetPreparationDimension dimension, bool storedWorksheet,
        bool mixedPrefixValues = false) {
        (string Name, string Xml)[] parts = [
            (PartNames[0], XmlDeclaration
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "</Types>"),
            (PartNames[1], XmlDeclaration
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"" + RelationshipNamespace + "/officeDocument\" Target=\"xl/workbook.xml\"/>"
                + "</Relationships>"),
            (PartNames[2], XmlDeclaration + "<workbook xmlns=\"" + SpreadsheetNamespace
                + "\" xmlns:r=\"" + RelationshipNamespace + "\"><sheets>"
                + "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"),
            (PartNames[3], XmlDeclaration
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"" + RelationshipNamespace + "/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "</Relationships>"),
            (WorksheetPartName, CreateWorksheet(dataRows, dimension, mixedPrefixValues))
        ];
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true, entryNameEncoding: Utf8)) {
            foreach ((string name, string xml) in parts) {
                CompressionLevel compression = storedWorksheet && name == WorksheetPartName
                    ? CompressionLevel.NoCompression : CompressionLevel.Optimal;
                ZipArchiveEntry entry = archive.CreateEntry(name, compression);
                entry.LastWriteTime = FixedTimestamp;
                entry.ExternalAttributes = 0;
                using Stream destination = entry.Open();
                destination.Write(Utf8.GetBytes(xml));
            }
        }
        return output.ToArray();
    }

    internal static void ValidatePackage(byte[] bytes, int dataRows,
        WorksheetPreparationDimension dimension, bool storedWorksheet, bool mixedPrefixValues = false) {
        using (var input = new MemoryStream(bytes, writable: false))
        using (var archive = new ZipArchive(input, ZipArchiveMode.Read)) {
            if (archive.Entries.Count != PartNames.Length
                || !archive.Entries.Select(entry => entry.FullName).ToHashSet(StringComparer.Ordinal).SetEquals(PartNames)) {
                throw new InvalidDataException("Preparation package entries differ.");
            }
            Console.WriteLine($"WorksheetPreparation package: rows={dataRows}; dimension={dimension}; "
                + $"storedWorksheet={storedWorksheet}; markup={(mixedPrefixValues ? WorksheetPreparationMarkup.MixedPrefixValues : WorksheetPreparationMarkup.Canonical)}; "
                + $"bytes={bytes.Length}; SHA256={Convert.ToHexString(SHA256.HashData(bytes))}.");
            foreach (string name in PartNames) {
                ZipArchiveEntry entry = archive.GetEntry(name) ?? throw new InvalidDataException($"Preparation part '{name}' is missing.");
                using Stream part = entry.Open();
                using var inflated = new MemoryStream();
                part.CopyTo(inflated);
                byte[] payload = inflated.ToArray();
                if (payload.LongLength != entry.Length)
                    throw new InvalidDataException($"Preparation part '{name}' length differs.");
                if (name == WorksheetPartName && (storedWorksheet
                    ? entry.CompressedLength != entry.Length : entry.CompressedLength >= entry.Length)) {
                    throw new InvalidDataException("Preparation worksheet ZIP storage differs.");
                }
                if (name == WorksheetPartName) ValidateWorksheetMarkup(payload, mixedPrefixValues);
                Console.WriteLine($"WorksheetPreparation part: name={name}; bytes={payload.Length}; "
                    + $"compressedBytes={entry.CompressedLength}; SHA256={Convert.ToHexString(SHA256.HashData(payload))}.");
            }
        }
        using var packageStream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(packageStream, isEditable: false);
        var errors = new OpenXmlValidator().Validate(document).Take(3).ToArray();
        if (errors.Length != 0) {
            throw new InvalidDataException("Preparation package is invalid Open XML: "
                + string.Join(" | ", errors.Select(error => error.Description)));
        }
        WorkbookPart workbook = document.WorkbookPart ?? throw new InvalidDataException("Preparation workbook is missing.");
        Workbook root = workbook.Workbook ?? throw new InvalidDataException("Preparation workbook root is missing.");
        Sheet[] sheets = root.Sheets?.Elements<Sheet>().ToArray() ?? [];
        if (sheets.Length != 1 || sheets[0].Name?.Value != "Data"
            || workbook.GetPartById(sheets[0].Id?.Value ?? string.Empty) is not WorksheetPart worksheetPart) {
            throw new InvalidDataException("Preparation worksheet relationship differs.");
        }
        Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidDataException("Preparation worksheet root is missing.");
        if (worksheet.GetFirstChild<SheetDimension>()?.Reference?.Value != DimensionReference(dataRows, dimension))
            throw new InvalidDataException("Preparation dimension declaration differs.");
        SheetData data = worksheet.GetFirstChild<SheetData>() ?? throw new InvalidDataException("Preparation rows are missing.");
        if (data.Elements<Row>().Count() != dataRows + 1)
            throw new InvalidDataException("Preparation package physical row count differs.");
        Cell[] header = data.Elements<Row>().First().Elements<Cell>().ToArray();
        Cell[] last = data.Elements<Row>().Last().Elements<Cell>().ToArray();
        string lastRow = (dataRows + 1).ToString(CultureInfo.InvariantCulture);
        if (header.Length != 2 || header[0].CellReference?.Value != "A1" || header[1].CellReference?.Value != "B1"
            || last.Length != 2 || last[0].CellReference?.Value != "A" + lastRow || last[1].CellReference?.Value != "C" + lastRow) {
            throw new InvalidDataException("Preparation header or late wider row coordinates differ.");
        }
    }

    private static void ValidateWorksheetMarkup(byte[] payload, bool mixedPrefixValues) {
        using var input = new MemoryStream(payload, writable: false);
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        if (reader.MoveToContent() != XmlNodeType.Element || reader.LocalName != "worksheet"
            || reader.Prefix.Length != 0 || reader.NamespaceURI != SpreadsheetNamespace
            || reader.LookupNamespace("p") != (mixedPrefixValues ? SpreadsheetNamespace : null)) {
            throw new InvalidDataException("Preparation worksheet root namespace differs.");
        }
        string? row = null;
        string? cell = null;
        int prefixedValues = 0;
        while (reader.Read()) {
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.LocalName == "row") row = reader.GetAttribute("r");
            if (reader.LocalName == "c") cell = reader.GetAttribute("r");
            if (reader.Prefix.Length == 0) continue;
            if (!mixedPrefixValues || reader.Prefix != "p" || reader.LocalName != "v"
                || reader.NamespaceURI != SpreadsheetNamespace || row != "2" || cell != "A2"
                || reader.ReadElementContentAsString() != "1") {
                throw new InvalidDataException("Preparation mixed prefix must be the same-namespace value in A2 only.");
            }
            prefixedValues++;
        }
        if (prefixedValues != (mixedPrefixValues ? 1 : 0))
            throw new InvalidDataException("Preparation mixed-prefix value count differs.");
        Console.WriteLine($"WorksheetPreparation markup: mixedPrefixValues={mixedPrefixValues}; prefixedValues={prefixedValues}.");
    }

    private static string CreateWorksheet(int dataRows, WorksheetPreparationDimension dimension,
        bool mixedPrefixValues = false) {
        var xml = new StringBuilder(checked(dataRows * 100 + 512));
        xml.Append(XmlDeclaration).Append("<worksheet xmlns=\"").Append(SpreadsheetNamespace).Append('"');
        if (mixedPrefixValues) xml.Append(" xmlns:p=\"").Append(SpreadsheetNamespace).Append('"');
        xml.Append('>');
        string? reference = DimensionReference(dataRows, dimension);
        if (reference != null) xml.Append("<dimension ref=\"").Append(reference).Append("\"/>");
        xml.Append("<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Id</t></is></c>"
            + "<c r=\"B1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>");
        for (int id = 1; id <= dataRows; id++) {
            string row = (id + 1).ToString(CultureInfo.InvariantCulture);
            string valueTag = mixedPrefixValues && id == 1 ? "p:v" : "v";
            xml.Append("<row r=\"").Append(row).Append("\"><c r=\"A").Append(row).Append("\"><").Append(valueTag).Append('>')
                .Append(id.ToString(CultureInfo.InvariantCulture)).Append("</").Append(valueTag).Append("></c><c r=\"")
                .Append(id == dataRows ? 'C' : 'B').Append(row).Append("\"><v>")
                .Append(Value(id).ToString("R", CultureInfo.InvariantCulture)).Append("</v></c></row>");
        }
        return xml.Append("</sheetData></worksheet>").ToString();
    }

    private static string? DimensionReference(int dataRows, WorksheetPreparationDimension dimension) => dimension switch {
        WorksheetPreparationDimension.Correct => "A1:C" + (dataRows + 1).ToString(CultureInfo.InvariantCulture),
        WorksheetPreparationDimension.Absent => null,
        WorksheetPreparationDimension.StaleNarrow => "A1:B" + (dataRows + 1).ToString(CultureInfo.InvariantCulture),
        _ => throw new ArgumentOutOfRangeException(nameof(dimension))
    };
}
