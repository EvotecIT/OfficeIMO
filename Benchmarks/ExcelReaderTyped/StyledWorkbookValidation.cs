#if OFFICEIMO_BENCHMARK_NEW_APIS
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    internal static class StyledWorkbookValidation {
        internal static void Validate(byte[] bytes, string engine, bool requireCoordinates) {
            using MemoryStream input = new MemoryStream(bytes, writable: false);
            using ZipArchive package = new ZipArchive(input, ZipArchiveMode.Read);
            WrittenWorkbookValidation.ValidateEntries(package);
            WrittenWorkbookValidation.ValidateRelationships(package);
            XDocument workbook = ReadXml(package.GetEntry("xl/workbook.xml")!);
            XElement[] sheets = workbook.Descendants().Where(element => element.Name.LocalName == "sheet").ToArray();
            if (sheets.Length != 1 || (string?)sheets[0].Attribute("name") != "S")
                throw new InvalidDataException("Styled output has incorrect sheet metadata.");
            XDocument styles = ReadXml(package.GetEntry("xl/styles.xml")!);
            XElement[] fonts = Group(styles, "fonts"), fills = Group(styles, "fills"), borders = Group(styles, "borders");
            XElement[] cellFormats = Group(styles, "cellXfs");
            foreach (XElement format in cellFormats) {
                Index(format, "fontId", fonts.Length);
                Index(format, "fillId", fills.Length);
                Index(format, "borderId", borders.Length);
            }
            int rows = 0, cells = 0, rowReferences = 0, cellReferences = 0;
            using (XmlReader reader = CreateReader(package.GetEntry("xl/worksheets/sheet1.xml")!.Open())) {
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row") {
                        if (cells != rows) throw new InvalidDataException("Styled output has an incorrect row width.");
                        continue;
                    }
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (reader.LocalName == "dimension") {
                        if (reader.GetAttribute("ref") != $"A1:A{StyledRowWriterBenchmarks.Rows}")
                            throw new InvalidDataException("Styled output has an incorrect dimension.");
                    } else if (reader.LocalName == "col") {
                        if (!int.TryParse(reader.GetAttribute("min"), out int first) || !int.TryParse(reader.GetAttribute("max"), out int last)
                            || first < 1 || last < first || last > 16_384)
                            throw new InvalidDataException("Styled output has an invalid column range.");
                        if (reader.GetAttribute("style") is string columnStyle) StyleIndex(columnStyle, cellFormats.Length);
                    } else if (reader.LocalName == "row") {
                        if (rows != cells || ++rows > StyledRowWriterBenchmarks.Rows)
                            throw new InvalidDataException("Styled output has missing or extra rows/cells.");
                        if (reader.GetAttribute("r") is string coordinate) {
                            rowReferences++;
                            if (coordinate != rows.ToString(CultureInfo.InvariantCulture))
                                throw new InvalidDataException("Styled output row order differs.");
                        }
                        if (reader.GetAttribute("customFormat") is not ("1" or "true"))
                            throw new InvalidDataException("Styled output omitted the real row default.");
                        RequireBold(reader.GetAttribute("s"), cellFormats, fonts);
                    } else if (reader.LocalName == "c") {
                        if (++cells != rows || reader.GetAttribute("t") is not (null or "n"))
                            throw new InvalidDataException("Styled output has an incorrect numeric cell or width.");
                        if (reader.GetAttribute("r") is string coordinate) {
                            cellReferences++;
                            if (coordinate != $"A{rows}") throw new InvalidDataException("Styled output has an incorrect cell coordinate.");
                        }
                        // Both row defaults and explicitly authored values must resolve to the bold font.
                        RequireBold(reader.GetAttribute("s"), cellFormats, fonts);
                        using XmlReader cell = reader.ReadSubtree();
                        string? value = null;
                        while (cell.Read())
                            if (cell.NodeType == XmlNodeType.Element && cell.LocalName == "v") value = cell.ReadElementContentAsString();
                        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) || number != rows - 1)
                            throw new InvalidDataException("Styled output numeric value differs.");
                    }
                }
            }
            if (rows != StyledRowWriterBenchmarks.Rows || cells != rows
                || (requireCoordinates && (rowReferences != rows || cellReferences != rows)))
                throw new InvalidDataException("Styled output count or coordinate policy differs.");
            WorkbookScanWorkload scan = new WorkbookScanWorkload();
            scan.Setup(bytes, ComparisonWorkbookFormat.Xlsx, rows, 1, row => [(double)row]);
            Console.WriteLine($"Validated {engine} styled output: rows={rows}, cells={cells}, boldRowDefaults={rows}, "
                + $"boldAuthoredCells={cells}, rowReferences={rowReferences}, cellReferences={cellReferences}, "
                + $"fonts={fonts.Length}, cellXfs={cellFormats.Length}, bytes={bytes.Length}, "
                + $"SHA256={Convert.ToHexString(SHA256.HashData(bytes))}, retainedOutputCapacity=32MiB, headers=false.");
        }

        private static void RequireBold(string? style, XElement[] formats, XElement[] fonts) {
            int index = StyleIndex(style, formats.Length);
            XElement font = fonts[Index(formats[index], "fontId", fonts.Length)];
            XElement? bold = font.Elements().SingleOrDefault(element => element.Name.LocalName == "b");
            if (bold == null || (string?)bold.Attribute("val") is "0" or "false")
                throw new InvalidDataException("Styled output does not resolve to a bold font.");
        }

        private static int StyleIndex(string? value, int length) => int.TryParse(value, NumberStyles.None,
            CultureInfo.InvariantCulture, out int index) && index >= 0 && index < length
            ? index : throw new InvalidDataException("Styled output references an invalid XF.");

        private static int Index(XElement element, string name, int length) => StyleIndex((string?)element.Attribute(name) ?? "0", length);

        private static XElement[] Group(XDocument document, string name) {
            XElement group = document.Root!.Elements().Single(element => element.Name.LocalName == name);
            XElement[] items = group.Elements().ToArray();
            if (!int.TryParse((string?)group.Attribute("count"), out int count) || count != items.Length || count == 0)
                throw new InvalidDataException("Styled output has an incorrect stylesheet collection count.");
            return items;
        }

        private static XDocument ReadXml(ZipArchiveEntry entry) {
            using XmlReader reader = CreateReader(entry.Open());
            return XDocument.Load(reader);
        }

        private static XmlReader CreateReader(Stream stream) => XmlReader.Create(stream, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, CloseInput = true,
        });
    }
}
#endif