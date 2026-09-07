using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Saves a reusable tagged DOCX template, then fills a separate request from structured values.</summary>
internal static class ServiceIntake {
    internal static void Create(string folder) {
        string templatePath = Path.Combine(folder, "template.docx");
        using (WordDocument template = WordDocument.Create(templatePath)) {
            template.Settings.FontFamily = "Carlito";
            template.Settings.FontSize = 11;
            template.AddParagraph("Service request").Style = WordParagraphStyles.Heading1;
            template.AddParagraph("NORTHWIND / INTAKE / EDITABLE TEMPLATE");
            AddField("Request identifier", "request-id");
            AddField("Requested by", "requester");
            AddField("Service", "service");
            AddField("Business need", "need");
            AddField("Required outcome", "outcome");
            template.AddParagraph("Triage guidance").Style = WordParagraphStyles.Heading2;
            template.AddParagraph("Confirm the owner, clarify acceptance criteria and agree the next update before scheduling work.");
            foreach (var paragraph in template.Paragraphs) {
                paragraph.FontFamily = "Carlito";
                paragraph.FontSize = paragraph.Style == WordParagraphStyles.Heading1 ? 26 : paragraph.Style == WordParagraphStyles.Heading2 ? 15 : 11;
                paragraph.LineSpacingAfterPoints = 8;
            }
            template.Save();

            void AddField(string label, string tag) {
                template.AddParagraph(label).Bold = true;
                template.AddStructuredDocumentTag("Enter " + label.ToLowerInvariant(), alias: label, tag: tag);
            }
        }
        using WordDocument request = WordDocument.Load(templatePath);
        var values = new Dictionary<string, string> {
            ["request-id"] = "SR-2048",
            ["requester"] = "Maya Ellis / Customer Operations",
            ["service"] = "Request portal",
            ["need"] = "The support team needs a weekly view of requests waiting for customer input.",
            ["outcome"] = "A report shows the owner, age and next action for each waiting request."
        };
        foreach (var value in values) {
            var control = request.GetStructuredDocumentTagByTag(value.Key)
                ?? throw new InvalidOperationException("Missing template field: " + value.Key);
            control.Text = value.Value;
        }
        request.Save(Path.Combine(folder, "example.docx"));
    }
}
