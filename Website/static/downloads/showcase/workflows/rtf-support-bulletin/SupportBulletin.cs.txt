using OfficeIMO.Rtf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a compact RTF support bulletin with formatted headings and structured troubleshooting steps.</summary>
internal static class SupportBulletin {
    internal static void Create(string folder) {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "Support bulletin: delayed request notifications";
        document.Info.Author = "OfficeIMO";
        int navy = document.AddColor(23, 54, 93), fill = document.AddColor(232, 238, 248);
        Heading("Delayed request notifications", 22);
        document.AddParagraph("SUPPORT BULLETIN / SB-014 / ILLUSTRATIVE SERVICE SCENARIO");
        Heading("Recognize the symptom");
        document.AddParagraph("A request appears in the portal, but its email acknowledgement arrives later than expected. Check the request record before asking the user to submit it again.");
        Heading("Check in this order");
        foreach (string step in new[] {
            "Confirm the request identifier and submission time.",
            "Check whether the notification is queued, delivered or rejected.",
            "Record the delivery result and route unresolved cases to the messaging owner."
        }) document.AddParagraph(step).SetList(kind: RtfListKind.Bullet).SetIndentation(leftTwips: 720, firstLineTwips: -360);
        Heading("Record enough evidence");
        var table = document.AddTable(4, 2); table.Rows[0].RepeatHeader = true; table.Rows[0].SetBackgroundColor(fill);
        string[,] rows = {
            { "Field", "Why it matters" },
            { "Request ID", "Finds the original service record" },
            { "Submission time", "Matches the notification event" },
            { "Delivery result", "Separates delay from rejection" }
        };
        for (int row = 0; row < 4; row++) for (int column = 0; column < 2; column++) table.Rows[row].Cells[column].AddParagraph(rows[row, column]);
        Heading("Close the loop");
        document.AddParagraph("Tell the user whether the request was received and when to expect the next update. Avoid including personal message content in the support record.");
        File.WriteAllText(Path.Combine(folder, "example.rtf"), document.ToRtf());

        void Heading(string text, double size = 14) {
            var paragraph = document.AddParagraph(text);
            paragraph.SpaceBeforeTwips = 160;
            paragraph.SpaceAfterTwips = 100;
            foreach (var run in paragraph.Runs) {
                run.FontSize = size;
                run.Bold = true;
                run.ForegroundColorIndex = navy;
            }
        }
    }
}
