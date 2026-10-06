using System.IO.Compression;
using System.Xml.Linq;

/// <summary>Independent stored-cache and preservation checks for the small report artifact lane.</summary>
internal static class ReportVerifier {
    internal static void Verify(string path) {
        using var zip = ZipFile.OpenRead(path);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XDocument Read(string name) { using var stream = zip.GetEntry(name)!.Open(); return XDocument.Load(stream); }
        var sheet = Read("xl/worksheets/sheet1.xml");
        if (Path.GetFileName(path).StartsWith("report-dates-", StringComparison.Ordinal)) {
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
