using OfficeIMO.Word;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Adds the running footer every Word workflow shares: the document title on the left and "Internal · Page N of M" on the right.</summary>
internal static class WorkflowFooter {
    internal static void Add(WordDocument document, string title) {
        if (document.Footer?.Default == null) document.AddHeadersAndFooters();
        WordSection section = document.Sections[0];
        int textWidth = (int)(section.PageSettings.Width!.Value - section.Margins.Left - section.Margins.Right);

        // One paragraph with a right tab at the text width keeps the line from splitting.
        WordParagraph footer = document.Footer!.Default!.AddParagraph(title);
        footer.AddTabStop(textWidth, WordTabAlignment.Right);
        footer.AddTab();
        footer.AddText("Internal · Page ");
        footer.AddField(WordFieldType.Page);
        footer.AddText(" of ");
        footer.AddField(WordFieldType.NumPages);
        foreach (WordParagraph run in footer.GetRuns()) {
            run.FontSize = 9; // FontSize takes whole points; the design asks for 8.5.
            run.ColorHex = "5B6575";
        }
    }
}
