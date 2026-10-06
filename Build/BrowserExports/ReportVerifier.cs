using System.IO.Compression;
using System.Xml.Linq;
using System.Text.Json;

/// <summary>Independent stored-cache and preservation checks for the small report artifact lane.</summary>
internal static class ReportVerifier {
    internal static void Verify(string path, string fixtureJson) {
        using var zip = ZipFile.OpenRead(path);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XDocument Read(string name) { using var stream = zip.GetEntry(name)!.Open(); return XDocument.Load(stream); }
        var sheet = Read("xl/worksheets/sheet1.xml");
        if (Path.GetFileName(path).StartsWith("report-regions-", StringComparison.Ordinal)) {
            string[] merges = sheet.Descendants(ns + "mergeCell").Select(m => (string)m.Attribute("ref")!).ToArray();
            if (!merges.SequenceEqual(new[] { "A1:B1", "A2:B2" })) throw new InvalidDataException("Title/group merges differ.");
            if ((string?)sheet.Descendants(ns + "pane").Single().Attribute("topLeftCell") != "A4") throw new InvalidDataException("Title freeze offset differs.");
            var table = Read("xl/tables/table1.xml");
            if ((string?)table.Root!.Attribute("ref") != "A3:B6") throw new InvalidDataException("Title table bounds differ.");
            var total = sheet.Descendants(ns + "c").Single(c => (string?)c.Attribute("r") == "B6");
            if ((string?)total.Element(ns + "f") != "SUBTOTAL(109,B4:B5)" || (double?)total.Element(ns + "v") != 20) throw new InvalidDataException("Title total differs.");
            var workbook = Read("xl/workbook.xml");
            if ((string?)workbook.Descendants(ns + "definedName").Single(n => (string?)n.Attribute("name") == "_xlnm.Print_Titles") != "'Report'!$2:$3") throw new InvalidDataException("Repeated headings include the title.");
            var regions = Read("xl/worksheets/sheet2.xml");
            if (!regions.Descendants(ns + "mergeCell").Select(m => (string)m.Attribute("ref")!).SequenceEqual(new[] { "A1:B2", "A3:B3" })) throw new InvalidDataException("Explicit merges differ.");
            if (Read("xl/tables/table2.xml").Root?.Attribute("ref")?.Value != "A2:B3") throw new InvalidDataException("Second report table differs.");
            using var corpus = JsonDocument.Parse(fixtureJson);
            var image = corpus.RootElement.GetProperty("cases").EnumerateArray().Single(c => c.GetProperty("name").GetString() == "report-table").GetProperty("sheets")[0].GetProperty("images")[0];
            byte[] expectedPng = Convert.FromBase64String(image.GetProperty("pngBase64").GetString()!);
            using var media = zip.Entries.Single(e => e.FullName.StartsWith("xl/media/", StringComparison.Ordinal)).Open();
            using var payload = new MemoryStream(); media.CopyTo(payload);
            if (!payload.ToArray().SequenceEqual(expectedPng)) throw new InvalidDataException("ChartForgeX PNG bytes differ.");
            if (!Read("xl/worksheets/sheet4.xml").Descendants(ns + "drawing").Any()) throw new InvalidDataException("Report chart drawing missing.");
        } else if (Path.GetFileName(path).StartsWith("report-dates-", StringComparison.Ordinal)) {
            double[] expected = { 92604, 2, 46302, 46301, 46303 };
            for (int i = 0; i < expected.Length; i++) {
                string cell = ((char)('A' + i)).ToString() + "4";
                var node = sheet.Descendants(ns + "c").Single(c => (string?)c.Attribute("r") == cell);
                if ((double?)node.Element(ns + "v") != expected[i]) throw new InvalidDataException("Report date cache differs: " + cell);
                if (i == 1 && (string?)node.Attribute("s") != "0") throw new InvalidDataException("Numeric count uses a date format.");
            }
        } else {
            if (!sheet.Descendants(ns + "hyperlink").Any(link => (string?)link.Attribute("ref") == "A3")) throw new InvalidDataException("Footer preservation link missing.");
            var overflow = Read("xl/worksheets/sheet2.xml");
            string text = string.Concat(overflow.Descendants(ns + "c").Where(c => ((string?)c.Attribute("r"))?.StartsWith("D", StringComparison.Ordinal) == true && (string?)c.Attribute("r") != "D1").SelectMany(c => c.Descendants(ns + "t")).Select(t => t.Value));
            if (text != new string('x', 100000)) throw new InvalidDataException("Preserved footer text differs.");
        }
    }
}
