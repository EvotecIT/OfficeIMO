using OfficeIMO.Word;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".docx");
try {
    using (WordDocument document = WordDocument.Create(path)) {
        document.AddParagraph("OfficeIMO NativeAOT Word marker for {{Customer.Name}}");
        IDictionary<string, object?> values = new Dictionary<string, object?> {
            ["Customer"] = new Dictionary<string, object?> { ["Name"] = "Northwind" }
        };
        WordTemplate.Apply(document, values).EnsureComplete();
        document.Save();
    }

    using WordDocument reopened = WordDocument.Load(path);
    if (!reopened.Paragraphs.Any(paragraph => paragraph.Text.Contains("NativeAOT Word marker for Northwind", StringComparison.Ordinal))) {
        throw new InvalidOperationException("The Word template binding round trip lost its marker paragraph.");
    }

    Console.WriteLine("PASS | Word template binding, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}
