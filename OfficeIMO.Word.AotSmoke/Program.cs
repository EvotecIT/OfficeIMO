using System.Text.RegularExpressions;
using OfficeIMO.Word;

PhysicalFontSmoke.Run();
StylesThemeSmoke.Run();

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".docx");
try {
    using (WordDocument document = WordDocument.Create(path)) {
        document.AddParagraph("OfficeIMO NativeAOT Word marker for {{Customer.Name}}");
        IDictionary<string, object?> values = new Dictionary<string, object?> {
            ["Customer"] = new Dictionary<string, object?> { ["Name"] = "Northwind" }
        };
        WordTemplate.Apply(document, values).EnsureComplete();
        WordParagraph order = document.AddParagraph("Order ");
        order.AddText("12-3").SetBold();
        order.AddText("4 ships");
        document.ReplaceAll(new Regex(@"(\d+)-(\d+)"), "$2/$1");
        document.Find("34/12").Ranges.Single().AddBookmark("Order");
        document.Save();
    }

    using WordDocument reopened = WordDocument.Load(path);
    if (!reopened.Paragraphs.Any(paragraph => paragraph.Text.Contains("NativeAOT Word marker for Northwind", StringComparison.Ordinal))) {
        throw new InvalidOperationException("The Word template binding round trip lost its marker paragraph.");
    }
    if (reopened.Bookmarks.Single().Text != "34/12") {
        throw new InvalidOperationException("The Word range replacement or spanning bookmark did not survive the round trip.");
    }

    Console.WriteLine("PASS | Word template binding, range replacement, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}
