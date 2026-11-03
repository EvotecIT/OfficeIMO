using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds a PDF reading list with descriptive paragraph and table hyperlinks.</summary>
internal static class LinkedResourceGuide {
    internal static void Create(string folder) {
        string repository = "https://github.com/EvotecIT/OfficeIMO";
        PdfTableCell[][] rows = {
            new[] { PdfTableCell.TextCell("Resource"), PdfTableCell.TextCell("Use it for") },
            new[] { PdfTableCell.TextCell("Source repository", repository), PdfTableCell.TextCell("Browse code, examples and package READMEs.") },
            new[] { PdfTableCell.TextCell("Published releases", repository + "/releases"), PdfTableCell.TextCell("Read release notes and inspect release assets.") },
            new[] { PdfTableCell.TextCell("Issue tracker", repository + "/issues"), PdfTableCell.TextCell("Search known problems before filing a reproducible report.") }
        };
        PdfDocument.Create(pdf => pdf.Content(content => content
            .Title("Document automation reading list")
            .Subtitle("A compact PDF reference for a team starting a document-generation workflow.")
            .H2("Start with a runnable example")
            .AddParagraph(p => {
                p.AddText("Open the ");
                p.AddLink("OfficeIMO examples directory", repository + "/tree/master/OfficeIMO.Examples");
                p.AddText(" and choose a format close to your intended output. Keep the generated file alongside its input data.");
            })
            .H2("Keep these references close")
            .AddTable(rows)
            .H2("When reporting a problem")
            .AddList(new[] {
                "Include the smallest input that reproduces the issue.",
                "State the package version, target framework and operating system.",
                "Describe the expected result and attach a safe sample of the actual output."
            })
            .AddParagraph(p => p.AddText("The resource names are clickable annotations. A static image shows their labels; open the PDF to follow them."))),
            new PdfCreateOptions { ShowPageNumbers = true })
            .Meta(title: "Document automation reading list", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
