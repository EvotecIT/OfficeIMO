using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Examples.Showcase;

/// <summary>
/// Generates compact, real-world source documents together with first-party PDF and image evidence
/// for the public website workflow catalog.
/// </summary>
internal static class RealWorldWorkflowShowcase {
    internal static void Example(string folderPath) {
        string output = Path.Combine(folderPath, "WorkflowShowcase");
        Directory.CreateDirectory(output);

        CreateWordDeliverySummary(output);
        CreateRtfChangeApproval(output);
        CreateMarkdownReleaseReadiness(output);
        CreateOpenDocumentHandover(output);
    }

    private static void CreateWordDeliverySummary(string output) {
        string sourcePath = Path.Combine(output, "customer-delivery-summary.docx");
        using WordDocument document = WordDocument.Create(sourcePath);
        document.BuiltinDocumentProperties.Title = "Customer delivery summary";
        document.BuiltinDocumentProperties.Creator = "OfficeIMO";

        document.AddParagraph("Customer delivery summary").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("A concise handover report generated from structured project data.");

        document.AddParagraph("Engagement overview").Style = WordParagraphStyles.Heading2;
        WordTable overview = document.AddTable(5, 2, WordTableStyle.TableGrid);
        SetWordRow(overview, 0, "Field", "Value");
        SetWordRow(overview, 1, "Customer", "Northwind Operations");
        SetWordRow(overview, 2, "Workstream", "Quarterly service review");
        SetWordRow(overview, 3, "Delivery status", "Ready for acceptance");
        SetWordRow(overview, 4, "Owner", "Customer Success");

        document.AddParagraph("Delivered outcomes").Style = WordParagraphStyles.Heading2;
        WordList outcomes = document.AddList(WordListStyle.Bulleted);
        outcomes.AddItem("Validated the production rollout and recovery path.");
        outcomes.AddItem("Documented ownership, support contacts, and acceptance criteria.");
        outcomes.AddItem("Prepared the next-quarter improvement backlog.");

        document.AddParagraph("Next steps").Style = WordParagraphStyles.Heading2;
        WordTable actions = document.AddTable(4, 3, WordTableStyle.TableGrid);
        SetWordRow(actions, 0, "Action", "Owner", "Status");
        SetWordRow(actions, 1, "Approve delivery notes", "Customer", "Pending");
        SetWordRow(actions, 2, "Publish runbook", "Delivery team", "Complete");
        SetWordRow(actions, 3, "Schedule health review", "Service owner", "Planned");

        document.Save();
        PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(
            new OfficeIMO.Word.Pdf.WordToPdfOptions().UseProfile(PdfExportProfile.Faithful));
        SavePdf(
            conversion,
            Path.Combine(output, "customer-delivery-summary.pdf"));
    }

    private static void CreateRtfChangeApproval(string output) {
        string sourcePath = Path.Combine(output, "change-approval-memo.rtf");
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "Change approval memo";
        document.Info.Author = "OfficeIMO";

        int accent = document.AddColor(35, 95, 180);
        int headerFill = document.AddColor(226, 236, 250);
        RtfStyle title = document.AddStyle(1, "Title");
        title.Bold = true;
        title.FontSize = 20;
        title.ForegroundColorIndex = accent;
        title.SpaceAfterTwips = 180;
        RtfStyle heading = document.AddStyle(2, "Heading 1");
        heading.Bold = true;
        heading.FontSize = 14;
        heading.ForegroundColorIndex = accent;
        heading.SpaceBeforeTwips = 180;
        heading.SpaceAfterTwips = 80;

        document.AddParagraph("Change approval memo").SetStyle(1);
        document.AddParagraph("Production maintenance window · Standard change · Ready for approval");
        document.AddParagraph("Decision summary").SetStyle(2);
        document.AddParagraph(
            "Approve a controlled configuration rollout with a tested rollback path and named validation owners.");

        RtfTable controls = document.AddTable(4, 3);
        controls.Rows[0].RepeatHeader = true;
        controls.Rows[0].SetBackgroundColor(headerFill);
        SetRtfRow(controls, 0, "Control", "Owner", "State");
        SetRtfRow(controls, 1, "Backup verified", "Operations", "Complete");
        SetRtfRow(controls, 2, "Rollback tested", "Engineering", "Complete");
        SetRtfRow(controls, 3, "Business approval", "Service owner", "Pending");

        document.AddParagraph("Execution checklist").SetStyle(2);
        AddRtfBullet(document, "Confirm monitoring and alert ownership.");
        AddRtfBullet(document, "Apply the change during the approved window.");
        AddRtfBullet(document, "Record validation results and close the change.");

        File.WriteAllText(sourcePath, document.ToRtf(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        PdfDocumentConversionResult conversion = document.ToPdfDocumentResult();
        SavePdf(
            conversion,
            Path.Combine(output, "change-approval-memo.pdf"));
    }

    private static void CreateMarkdownReleaseReadiness(string output) {
        string sourcePath = Path.Combine(output, "release-readiness-report.md");
        MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
            .WithColorScheme(MarkdownColorSchemeKind.Blue);
        MarkdownDoc document = MarkdownDoc.Create()
            .H1("Release readiness report")
            .P("A portable release decision record generated from pipeline and owner data.")
            .H2("Readiness summary")
            .Table(table => table
                .Headers("Gate", "Owner", "Result")
                .Row("Automated checks", "Engineering", "Passed")
                .Row("Package inspection", "Release manager", "Passed")
                .Row("Production approval", "Service owner", "Pending"))
            .H2("Release checklist")
            .Ul(list => list
                .Item("Versioned artifacts are reproducible.")
                .Item("Upgrade and rollback notes are published.")
                .Item("Monitoring ownership is confirmed."))
            .Callout(
                "warning",
                "Decision required",
                "Production publication starts only after the service owner records approval.");

        File.WriteAllText(sourcePath, document.ToMarkdown(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        MarkdownToPdfOptions options = new MarkdownToPdfOptions {
            Theme = theme
        }.UseProfile(PdfExportProfile.Faithful);
        PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options);
        SavePdf(
            conversion,
            Path.Combine(output, "release-readiness-report.pdf"));
    }

    private static void CreateOpenDocumentHandover(string output) {
        string sourcePath = Path.Combine(output, "project-handover-brief.odt");
        OdtDocument document = OdtDocument.Create();
        document.AddHeading("Project handover brief", 1);
        document.AddParagraph(
            "A vendor-neutral handover document for teams that standardize on OpenDocument.");
        document.AddHeading("Delivery package", 2);
        OdtTable package = document.AddTable(4, 3, "DeliveryPackage");
        SetOdtRow(package, 0, "Deliverable", "Owner", "Status");
        SetOdtRow(package, 1, "Operational runbook", "Operations", "Complete");
        SetOdtRow(package, 2, "Support contacts", "Service owner", "Complete");
        SetOdtRow(package, 3, "Follow-up backlog", "Product team", "Accepted");
        document.AddHeading("Handover notes", 2);
        document.AddParagraph(
            "The receiving team owns production monitoring from acceptance. The delivery team remains available for the agreed transition period.");
        document.AddParagraph(
            "Next review: confirm service health, unresolved risks, and the first backlog milestone.");
        document.Save(sourcePath);

        PdfDocumentConversionResult conversion = document.ToPdfDocumentResult();
        SavePdf(
            conversion,
            Path.Combine(output, "project-handover-brief.pdf"));
    }

    private static void SavePdf(PdfDocumentConversionResult conversion, string pdfPath) {
        conversion.Save(pdfPath);
        Console.WriteLine($"✓ PDF: {pdfPath}");
    }

    private static void SetWordRow(WordTable table, int row, params string[] values) {
        for (int column = 0; column < values.Length; column++) {
            table.Rows[row].Cells[column].Paragraphs[0].Text = values[column];
        }
    }

    private static void SetRtfRow(RtfTable table, int row, params string[] values) {
        for (int column = 0; column < values.Length; column++) {
            table.Rows[row].Cells[column].AddParagraph(values[column]);
        }
    }

    private static void AddRtfBullet(RtfDocument document, string text) {
        document.AddParagraph(text)
            .SetList(kind: RtfListKind.Bullet)
            .SetIndentation(leftTwips: 720, firstLineTwips: -360);
    }

    private static void SetOdtRow(OdtTable table, int row, params string[] values) {
        for (int column = 0; column < values.Length; column++) {
            table.Cell(row, column).Text = values[column];
        }
    }
}
