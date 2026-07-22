using OfficeIMO.Word;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".docx");
try {
    using (WordDocument document = WordDocument.Create(path)) {
        document.AddParagraph("OfficeIMO NativeAOT Word marker");
        document.Save();
    }

    using WordDocument reopened = WordDocument.Load(path);
    if (!reopened.Paragraphs.Any(paragraph => paragraph.Text.Contains("NativeAOT Word marker", StringComparison.Ordinal))) {
        throw new InvalidOperationException("The Word round trip lost its marker paragraph.");
    }

    Console.WriteLine("PASS | Word create, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}
