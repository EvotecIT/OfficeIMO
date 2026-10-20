using System.IO.Compression;
using System.Xml.Linq;

/// <summary>Read exact merged headings, selected typed cells and multirow footers independently of the browser writer.</summary>
internal static class StructuredTableVerifier {
    internal static object VerifyCanopy(string path, string mode) {
        void Require(bool value, string message) { if (!value) throw new InvalidDataException(Path.GetFileName(path) + ": " + message); }
        if (path.EndsWith(".csv", StringComparison.Ordinal)) {
            string rows = mode == "raw" ? "0001,12.5,2026-01-01T00:00:00Z\r\n0003,3,2026-01-03T00:00:00Z\r\n" : "Asset 0001,12.50 USD,1 Jan 2026\r\nAsset 0003,3.00 USD,3 Jan 2026\r\n";
            Require(File.ReadAllText(path) == "ID,Metrics,\r\n,Amount,Seen\r\n" + rows + "Totals,15.5,Approved\r\n,Finance,\r\n", "Raw/display structured CSV differs.");
        } else {
            WorkbookVerifier.Verify(path);
            using var zip = ZipFile.OpenRead(path);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XDocument Read(string name) { using var stream = zip.GetEntry(name)!.Open(); return XDocument.Load(stream); }
            var sheet = Read("xl/worksheets/sheet1.xml");
            var styles = Read("xl/styles.xml");
            var cells = sheet.Descendants(ns + "c").ToDictionary(cell => (string)cell.Attribute("r")!);
            string Text(string coordinate) => string.Concat(cells[coordinate].Descendants(ns + "t").Select(node => node.Value));
            Require(sheet.Descendants(ns + "mergeCell").Select(cell => (string)cell.Attribute("ref")!).ToHashSet().SetEquals(new[] { "A1:A2", "B1:C1", "A5:A6", "B6:C6" }), "Heading/footer rectangles differ.");
            Require(Text("A3") == (mode == "raw" ? "0001" : "Asset 0001") && Text("A4") == (mode == "raw" ? "0003" : "Asset 0003"), "Literal IDs/display labels differ.");
            if (mode == "raw") Require((double?)cells["B3"].Element(ns + "v") == 12.5 && (double?)cells["B4"].Element(ns + "v") == 3 && (double?)cells["C3"].Element(ns + "v") == 46023 && (double?)cells["C4"].Element(ns + "v") == 46025, "Raw numeric/date types differ.");
            else Require(Text("B3") == "12.50 USD" && Text("B4") == "3.00 USD" && Text("C3") == "1 Jan 2026" && Text("C4") == "3 Jan 2026", "Display capture differs.");
            var formats = styles.Descendants(ns + "cellXfs").Single().Elements().ToArray();
            var fills = styles.Descendants(ns + "fills").Single().Elements().ToArray();
            for (int row = 3; row <= 4; row++) foreach (string column in new[] { "A", "B", "C" }) {
                var cell = cells[column + row];
                var format = formats[(int?)cell.Attribute("s") ?? 0];
                var fill = fills[(int?)format.Attribute("fillId") ?? 0];
                Require(fill.Descendants(ns + "fgColor").Single().Attribute("rgb")?.Value == (row == 3 ? "FFFCE4D6" : "FFE2F0D9"), "Explicit named semantic tone was lost.");
            }
            Require(Text("A5") == "Totals" && Text("B6") == "Finance", "Footer values differ.");
        }
        return new { file = Path.GetFileName(path), rows = 2, columns = 3, values = mode, structured = true, namedTones = true };
    }

    internal static void Verify(string path) {
        using var zip = ZipFile.OpenRead(path);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XDocument Read(string name) { using var stream = zip.GetEntry(name)!.Open(); return XDocument.Load(stream); }
        var sheet = Read("xl/worksheets/sheet1.xml");
        var cells = sheet.Descendants(ns + "c").ToDictionary(cell => (string)cell.Attribute("r")!);
        string Text(string coordinate) => string.Concat(cells[coordinate].Descendants(ns + "t").Select(node => node.Value));
        void Require(bool value, string message) { if (!value) throw new InvalidDataException(Path.GetFileName(path) + ": " + message); }
        Require(sheet.Descendants(ns + "mergeCell").Select(cell => (string)cell.Attribute("ref")!).ToHashSet().SetEquals(new[] { "A1:A2", "B1:C1", "A5:A6", "B6:C6" }), "Heading/footer rectangles differ.");
        Require(Text("A1") == "ID" && Text("B1") == "Metrics" && Text("B2") == "Amount" && Text("C2") == "Seen", "Heading labels differ.");
        Require(Text("A3") == "0003" && Text("A4") == "0001" && (string?)cells["A3"].Attribute("t") == "inlineStr", "Selected identifier/order or literal type differs.");
        Require((double?)cells["B3"].Element(ns + "v") == 3 && (double?)cells["B4"].Element(ns + "v") == 12.5 && cells["B3"].Attribute("t") is null, "Numeric values differ.");
        Require((double?)cells["C3"].Element(ns + "v") == 46025 && (double?)cells["C4"].Element(ns + "v") == 46023, "Typed UTC dates differ.");
        Require(Text("A5") == "Totals" && Text("B5") == "15.5" && Text("C5") == "Approved" && Text("B6") == "Finance", "Multirow footer values differ.");
        Require(!sheet.Descendants(ns + "autoFilter").Any(), "Explicitly disabled auto filter returned.");
        Require((string?)sheet.Descendants(ns + "pane").Single().Attribute("topLeftCell") == "A3", "Frozen heading rows differ.");
        var workbook = Read("xl/workbook.xml");
        Require(workbook.Descendants(ns + "definedName").Any(node => (string?)node.Attribute("name") == "_xlnm.Print_Titles" && node.Value == "'Data'!$1:$2"), "Print repetition differs.");
        var styles = Read("xl/styles.xml");
        Require(styles.Descendants(ns + "numFmt").Any(node => (string?)node.Attribute("formatCode") == "yyyy-mm-dd"), "Typed date format missing.");
    }
}
