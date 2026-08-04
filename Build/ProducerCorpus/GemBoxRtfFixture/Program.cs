using GemBox.Document;
using GemBox.Document.Tables;

ComponentInfo.SetLicense("FREE-LIMITED-KEY");

string outputPath = args.Length == 1
    ? Path.GetFullPath(args[0])
    : throw new ArgumentException("Expected output RTF path.");

var document = new DocumentModel();
var section = new Section(document);
document.Sections.Add(section);

section.Blocks.Add(new Paragraph(document,
    new Run(document, "Commercial library RTF fixture") { CharacterFormat = { Bold = true } }));
section.Blocks.Add(new Paragraph(document,
    "Generated from original MIT-licensed test content with GemBox.Document 2026.8.102 in free-limited mode."));
section.Blocks.Add(new Paragraph(document,
    new Run(document, "Unicode: Zażółć gęślą jaźń — résumé")));

var table = new Table(document,
    new TableRow(document,
        new TableCell(document, new Paragraph(document, "Route")),
        new TableCell(document, new Paragraph(document, "Status"))),
    new TableRow(document,
        new TableCell(document, new Paragraph(document, "RTF to HTML")),
        new TableCell(document, new Paragraph(document, "Verified"))));
section.Blocks.Add(table);

document.Save(outputPath);
