using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds a PDF reading list with descriptive paragraph and table hyperlinks.</summary>
internal static class LinkedResourceGuide {
    internal static void Create(string folder) {
        PdfColor navy = PdfColor.FromRgb(23, 54, 93);
        string repository = "https://github.com/EvotecIT/OfficeIMO";
        string[][] rows = {
            new[] { "Resource", "Use it for" },
            new[] { "Source repository", "Browse code, examples and package READMEs." },
            new[] { "Published releases", "Read release notes and inspect release assets." },
            new[] { "Issue tracker", "Search known problems before filing a reproducible report." }
        };
        var links = new Dictionary<(int Row, int Col), string> {
            [(1, 0)] = repository,
            [(2, 0)] = repository + "/releases",
            [(3, 0)] = repository + "/issues"
        };
        PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("Document automation reading list", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("A compact PDF reference for a team starting a document-generation workflow."))
            .H2("Start with a runnable example", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("Open the ").Link("OfficeIMO examples directory", repository + "/tree/master/OfficeIMO.Examples")
                .Text(" and choose a format close to your intended output. Keep the generated file alongside its input data."))
            .H2("Keep these references close", PdfAlign.Left, navy)
            .TableWithLinks(rows, links, style: new PdfTableStyle {
                HeaderFill = navy, HeaderTextColor = PdfColor.White, HeaderRowCount = 1,
                RowStripeFill = PdfColor.FromRgb(241, 245, 249), CellPaddingX = 9, CellPaddingY = 11
            })
            .H2("When reporting a problem", PdfAlign.Left, navy)
            .Bullets(new[] {
                "Include the smallest input that reproduces the issue.",
                "State the package version, target framework and operating system.",
                "Describe the expected result and attach a safe sample of the actual output."
            })
            .Paragraph(p => p.Text("The resource names are clickable annotations. A static image shows their labels; open the PDF to follow them."))),
            new PdfOptions { DefaultFontSize = 11 })
            .Meta(title: "Document automation reading list", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
