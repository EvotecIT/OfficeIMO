using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class WordRtfWorkflowTests {
    [Fact]
    public void MailMerge_Routes_Through_Word_And_Returns_Combined_Result() {
        RtfDocument template = RtfDocument.Create();
        template.AddParagraph("Dear ").AddField("MERGEFIELD Name").AddText("recipient");

        RtfWordWorkflowResult<WordMailMergeTemplateInspection> result = template.MailMergeResult(
            new Dictionary<string, string> { ["Name"] = "Ada" });

        Assert.True(result.WorkflowResult.IsValid);
        Assert.Contains("Name", result.WorkflowResult.MergeFieldNames);
        Assert.Contains("Dear Ada", PlainText(result.Document), StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordMailMergeCompleted");
    }

    [Fact]
    public void FindAndReplace_Routes_Through_Word_Cross_Run_Engine() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Hello ").SetBold();
        paragraph.AddText("World").SetItalic();

        RtfWordWorkflowResult<int> result = document.FindAndReplaceResult("Hello World", "Goodbye");

        Assert.Equal(1, result.WorkflowResult);
        Assert.Contains("Goodbye", PlainText(result.Document), StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateFields_Returns_Word_Field_Report_And_Updated_Rtf() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "Quarterly report";
        document.AddParagraph().AddField("TITLE").AddText("old title");

        RtfWordWorkflowResult<WordFieldUpdateReport> result = document.UpdateFieldsResult();

        Assert.Equal(1, result.WorkflowResult.UpdatedCount);
        Assert.Contains("Quarterly report", PlainText(result.Document), StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordFieldUpdateIncomplete");
    }

    [Fact]
    public void Merge_Uses_Word_Merge_Engine_And_Returns_Rtf() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddParagraph("First");
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Second");

        RtfWordWorkflowResult<int> result = destination.MergeResult(source);

        Assert.Equal(1, result.WorkflowResult);
        Assert.Contains("First", PlainText(result.Document), StringComparison.Ordinal);
        Assert.Contains("Second", PlainText(result.Document), StringComparison.Ordinal);
    }

    [Fact]
    public void Compare_Uses_Word_Structural_Comparison_Without_Temporary_Files() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Before");
        RtfDocument target = RtfDocument.Create();
        target.AddParagraph("After");

        RtfWordWorkflowResult<WordComparisonResult> result = source.CompareResult(target);

        Assert.True(result.WorkflowResult.HasChanges);
        Assert.NotEmpty(result.WorkflowResult.Findings);
        Assert.Equal("source.rtf", result.WorkflowResult.SourcePath);
        Assert.Equal("target.rtf", result.WorkflowResult.TargetPath);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordCompareCompleted");
    }

    [Fact]
    public void Read_Result_Workflow_Retains_Core_Diagnostics() {
        const string rtf = @"{\rtf1\ansi{\*\vendorprivate hidden}\pard Hello\par}";
        RtfReadResult read = RtfDocument.Read(rtf);

        RtfWordWorkflowResult<int> result = read.FindAndReplaceResult("Hello", "Hi");

        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RTF101" && diagnostic.SourcePath == "source");
        Assert.Contains("Hi", PlainText(result.Document), StringComparison.Ordinal);
    }

    private static string PlainText(RtfDocument document) =>
        string.Join(Environment.NewLine, document.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
}
