using OfficeIMO.Word;

using var document = WordDocument.Create("Example.docx");
document.AddParagraph("Hello from OfficeIMO");
document.Save();
