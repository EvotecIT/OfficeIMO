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

// Page summaries are newer than the pinned published package, so its consumer run skips them.
#if !OFFICEIMO_DOCUMENTATION_PUBLISHED_PACKAGE
#region PageSubtotals
var invoiceStyle = PdfTableStyles.Minimal();
invoiceStyle.FooterRowCount = 1;
var pageSummary = invoiceStyle.PageSummary = new PdfTablePageSummary();
PdfTableAggregate amount = pageSummary.Sum("amount", column: 2, format: "N2");
pageSummary.AddContinuationHeaderRow("Brought forward", "", amount.BroughtForwardToken);
pageSummary.AddPageFooterRow("Carried forward", "", amount.RunningToken);

var invoiceLines = new List<string[]> { new[] { "Item", "Description", "Amount" } };
for (int line = 1; line <= 120; line++) invoiceLines.Add(new[] { line.ToString(), "Consulting hour", "125.00" });
invoiceLines.Add(new[] { "Total", "", amount.RunningToken });

PdfDocument.Create(pdf => pdf.Content(content => content.AddTable(invoiceLines, style: invoiceStyle)))
    .Save("Invoice-lines.pdf");
#endregion
#endif

// Reopen every saved file and check the content each example promises.
Require(result.Pages.Count == 1 && result.Pages[0].Tables.Count == 1, "The created report must read back as one page with one table");
Require(result.ExtractMarkdown().Contains("Rotate certificates"), "The read Markdown must contain the table rows");

PdfDocumentReadResult pack = PdfDocument.Load("Review-pack.pdf").Read();
Require(pack.Pages.Count == 2 && pack.Text.Contains("Appendix"), "The combined pack must contain both source documents");
PdfDocumentInfo packInfo = PdfDocument.Load("Review-pack.pdf").Inspect();
Require(packInfo.Pages[0].RotationDegrees == 0 && packInfo.Pages[1].RotationDegrees == 90, "Only page 2 of the pack must be rotated");
Require(parts.Count == 2, "Split must save one file per page");
Require(PdfDocument.Load("Review-pack-1.pdf").Read().Text.Contains("Rotate certificates"), "The first split file must hold the report page");
PdfDocument secondPart = PdfDocument.Load("Review-pack-2.pdf");
Require(secondPart.Read().Text.Contains("Backup schedule") && secondPart.Inspect().Pages[0].RotationDegrees == 90, "The second split file must hold the rotated appendix page");

PdfDocument filled = PdfDocument.Load("Application-filled.pdf");
Require(filled.Read().Text.Contains("Adele Vance") && filled.Inspect().FormFields.Count == 0, "Flattened values must become page text");

Require(protectedPdf.IsEncrypted, "The protected copy must be encrypted");
PdfDocument reopened = PdfDocument.Load("Service-review-protected.pdf", new PdfLoadOptions { Password = "open-password" });
Require(reopened.Read().Text.Contains("Service review"), "The protected copy must open with its password");

byte[] png = File.ReadAllBytes("Service-review-page-1.png");
Require(preview.Succeeded && png.Length > 8 && png[0] == 0x89 && png[1] == (byte)'P', "The rendered page must be a PNG");

#if !OFFICEIMO_DOCUMENTATION_PUBLISHED_PACKAGE
string[] invoicePages = PdfDocument.Load("Invoice-lines.pdf").Pages.Split().Select(page => page.Read().Text).ToArray();
Require(invoicePages.Length > 1, "The invoice table must continue across pages");
for (int page = 0; page < invoicePages.Length - 1; page++) {
    Require(invoicePages[page].Contains("Carried forward"), "Every continued page must close with the carried-forward amount");
    Require(invoicePages[page + 1].Contains("Brought forward"), "Every continuation page must open with the brought-forward amount");
}
Require(invoicePages[invoicePages.Length - 1].Contains("15,000.00"), "The final page must show the invoice total");
#endif

Console.WriteLine("PDF README examples reopened and matched their content.");

static void Require(bool condition, string message) {
    if (!condition) throw new InvalidOperationException(message);
}
