using OfficeIMO;
using OfficeIMO.Word;

internal static class HiddenDataSmoke {
    internal static void Run() {
        using WordDocument document = WordDocument.Create();
        document.Properties.Creator = "NativeAOT Author";
        document.AddParagraph("Reviewed").AddComment("NativeAOT Reviewer", "NR", "Check");
        document.AddTable(3, 2);
        OfficeHiddenDataReport report = document.RemoveHiddenData();
        if (!report.Changed || document.InspectHiddenData().GetFindings(OfficeHiddenDataKind.Comments).Count != 0) {
            throw new InvalidOperationException("Word hidden-data removal did not remove the comment.");
        }
        OfficeAccessibilityReport accessibility = document.FixAccessibility(new WordAccessibilityFixOptions { Title = "Smoke", Language = "en-US", MarkTableHeaderRows = true });
        if (accessibility.Findings.Any(finding => finding.Code == OfficeDiagnosticCodes.AccessibilityMissingTableHeader)) {
            throw new InvalidOperationException("Word accessibility fix did not mark the table header row.");
        }
        Console.WriteLine("PASS | Word hidden-data removal and accessibility fixes");
    }
}
