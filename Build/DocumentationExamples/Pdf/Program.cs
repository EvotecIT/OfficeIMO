#region Create
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Service review")
        .AddParagraph("All scheduled checks passed. Two follow-ups remain open.")
        .AddTable(new[] {
            new[] { "Task", "Owner", "Due" },
            new[] { "Rotate certificates", "Platform", "4 Nov" } })))
    .Meta(title: "Service review", author: "Operations")
    .Save("Service-review.pdf");
#endregion

PdfDocument.Create(pdf => pdf.Content(content => content.H1("Appendix").AddParagraph("Backup schedule.")))
    .Save("Appendix.pdf");
PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Application")
        .TextField("Applicant.Name")
        .TextField("Approval.Status")))
    .Save("Application.pdf");

#region Read
PdfDocumentReadResult result = PdfDocument.Load("Service-review.pdf").Read();
foreach (PdfLogicalPage page in result.Pages) {
    Console.WriteLine($"Page {page.PageNumber}: {page.Paragraphs.Count} paragraphs, {page.Tables.Count} tables");
}
Console.WriteLine(result.ExtractMarkdown());
#endregion

#region Combine
PdfDocument.Load("Service-review.pdf")
    .MergeWith("Appendix.pdf")
    .Pages.Rotate(90, "2")
    .Save("Review-pack.pdf");

IReadOnlyList<PdfDocument> parts = PdfDocument.Load("Review-pack.pdf").Pages.Split();
for (int i = 0; i < parts.Count; i++) parts[i].Save($"Review-pack-{i + 1}.pdf");
#endregion

#region FillForm
PdfDocument.Load("Application.pdf")
    .Forms.FillAndFlatten(new Dictionary<string, string> {
        ["Applicant.Name"] = "Adele Vance",
        ["Approval.Status"] = "Approved"
    })
    .Save("Application-filled.pdf");
#endregion

#region Protect
PdfSecurityMutationResult protectedPdf = PdfDocument.Load("Service-review.pdf")
    .Security.Encrypt(new PdfStandardEncryptionOptions("open-password") {
        OwnerPassword = "owner-password"
    });
File.WriteAllBytes("Service-review-protected.pdf", protectedPdf.Pdf);
#endregion

#region Render
PdfPageRenderResult preview = PdfDocument.Load("Service-review.pdf")
    .Render.Pages("1", new PdfPageRenderOptions { Dpi = 144 })[0];
File.WriteAllBytes("Service-review-page-1.png", preview.Bytes!);
#endregion

// Reopen every saved file and check the content each example promises.
Require(result.Pages.Count == 1 && result.Pages[0].Tables.Count == 1, "The created report must read back as one page with one table");
Require(result.ExtractMarkdown().Contains("Rotate certificates"), "The read Markdown must contain the table rows");

PdfDocumentReadResult pack = PdfDocument.Load("Review-pack.pdf").Read();
Require(pack.Pages.Count == 2 && pack.Text.Contains("Appendix"), "The combined pack must contain both source documents");
Require(parts.Count == 2 && PdfDocument.Load("Review-pack-2.pdf").Read().Text.Contains("Backup schedule"), "Split must save one file per page");

PdfDocument filled = PdfDocument.Load("Application-filled.pdf");
Require(filled.Read().Text.Contains("Adele Vance") && filled.Inspect().FormFields.Count == 0, "Flattened values must become page text");

Require(protectedPdf.IsEncrypted, "The protected copy must be encrypted");
PdfDocument reopened = PdfDocument.Load("Service-review-protected.pdf", new PdfLoadOptions { Password = "open-password" });
Require(reopened.Read().Text.Contains("Service review"), "The protected copy must open with its password");

byte[] png = File.ReadAllBytes("Service-review-page-1.png");
Require(preview.Succeeded && png.Length > 8 && png[0] == 0x89 && png[1] == (byte)'P', "The rendered page must be a PNG");

Console.WriteLine("PDF README examples reopened and matched their content.");

static void Require(bool condition, string message) {
    if (!condition) throw new InvalidOperationException(message);
}
