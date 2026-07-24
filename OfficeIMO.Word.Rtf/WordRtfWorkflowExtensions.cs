using System.Collections.Generic;

namespace OfficeIMO.Word.Rtf;

/// <summary>Thin RTF workflow routes over the reusable OfficeIMO.Word engines.</summary>
public static class WordRtfWorkflowExtensions {
    /// <summary>Executes Word mail merge against an RTF template and converts the result back to RTF.</summary>
    public static RtfWordWorkflowResult<WordMailMergeTemplateInspection> MailMergeResult(
        this RtfDocument template,
        IDictionary<string, string> values,
        bool removeFields = true) {
        if (template == null) throw new ArgumentNullException(nameof(template));
        if (values == null) throw new ArgumentNullException(nameof(values));

        RtfConversionResult<WordDocument> input = template.ToWordDocumentResult();
        using WordDocument word = input.Value;
        WordMailMergeTemplateInspection inspection = WordMailMerge.InspectTemplate(word, values.Keys);
        WordMailMerge.Execute(word, values, removeFields);
        return Complete(word, inspection, input.Report, "RtfWordMailMergeCompleted", "Mail merge was executed by OfficeIMO.Word.");
    }

    /// <summary>Executes Word mail merge from a parsed RTF result, retaining parser diagnostics.</summary>
    public static RtfWordWorkflowResult<WordMailMergeTemplateInspection> MailMergeResult(
        this RtfReadResult template,
        IDictionary<string, string> values,
        bool removeFields = true) {
        if (template == null) throw new ArgumentNullException(nameof(template));
        RtfWordWorkflowResult<WordMailMergeTemplateInspection> result = template.Document.MailMergeResult(values, removeFields);
        result.Report.AddReadDiagnostics(template.Diagnostics, "template");
        return result;
    }

    /// <summary>Runs the OfficeIMO.Word cross-run find/replace engine and returns normalized RTF.</summary>
    public static RtfWordWorkflowResult<int> FindAndReplaceResult(
        this RtfDocument document,
        string textToFind,
        string textToReplace,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfConversionResult<WordDocument> input = document.ToWordDocumentResult();
        using WordDocument word = input.Value;
        int replacements = word.FindAndReplace(textToFind, textToReplace, comparison);
        return Complete(word, replacements, input.Report, "RtfWordFindReplaceCompleted", "Find and replace was executed by OfficeIMO.Word.");
    }

    /// <summary>Runs find/replace from a parsed RTF result, retaining parser diagnostics.</summary>
    public static RtfWordWorkflowResult<int> FindAndReplaceResult(
        this RtfReadResult document,
        string textToFind,
        string textToReplace,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfWordWorkflowResult<int> result = document.Document.FindAndReplaceResult(textToFind, textToReplace, comparison);
        result.Report.AddReadDiagnostics(document.Diagnostics, "source");
        return result;
    }

    /// <summary>Updates deterministic fields through OfficeIMO.Word and returns its detailed update report.</summary>
    public static RtfWordWorkflowResult<WordFieldUpdateReport> UpdateFieldsResult(
        this RtfDocument document,
        WordFieldUpdateOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfConversionResult<WordDocument> input = document.ToWordDocumentResult();
        using WordDocument word = input.Value;
        WordFieldUpdateReport workflow = word.UpdateFieldsAndGetReport(options ?? WordFieldUpdateOptions.Default);
        RtfWordWorkflowResult<WordFieldUpdateReport> result = Complete(word, workflow, input.Report, "RtfWordFieldsUpdated", "Fields were evaluated by OfficeIMO.Word.");
        if (workflow.UnsupportedCount > 0 || workflow.ParseErrorCount > 0) {
            result.Report.Add(
                RtfConversionSeverity.Warning,
                "RtfWordFieldUpdateIncomplete",
                "One or more Word fields could not be evaluated.",
                RtfConversionAction.Omitted,
                feature: "field",
                count: workflow.UnsupportedCount + workflow.ParseErrorCount);
        }

        return result;
    }

    /// <summary>Updates fields from a parsed RTF result, retaining parser diagnostics.</summary>
    public static RtfWordWorkflowResult<WordFieldUpdateReport> UpdateFieldsResult(
        this RtfReadResult document,
        WordFieldUpdateOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfWordWorkflowResult<WordFieldUpdateReport> result = document.Document.UpdateFieldsResult(options);
        result.Report.AddReadDiagnostics(document.Diagnostics, "source");
        return result;
    }

    /// <summary>Appends another RTF document through the OfficeIMO.Word merge engine.</summary>
    public static RtfWordWorkflowResult<int> MergeResult(this RtfDocument document, RtfDocument source) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (source == null) throw new ArgumentNullException(nameof(source));
        RtfConversionResult<WordDocument> destinationConversion = document.ToWordDocumentResult();
        RtfConversionResult<WordDocument> sourceConversion = source.ToWordDocumentResult();
        using WordDocument destinationWord = destinationConversion.Value;
        using WordDocument sourceWord = sourceConversion.Value;
        destinationWord.AppendDocument(sourceWord);
        var report = new RtfConversionReport();
        report.Merge(destinationConversion.Report);
        report.Merge(sourceConversion.Report);
        return Complete(destinationWord, 1, report, "RtfWordMergeCompleted", "One RTF document was appended by OfficeIMO.Word.");
    }

    /// <summary>Compares two RTF documents using the machine-readable OfficeIMO.Word comparison engine.</summary>
    public static RtfWordWorkflowResult<WordComparisonResult> CompareResult(
        this RtfDocument source,
        RtfDocument target,
        WordComparisonOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (target == null) throw new ArgumentNullException(nameof(target));
        RtfConversionResult<WordDocument> sourceConversion = source.ToWordDocumentResult();
        RtfConversionResult<WordDocument> targetConversion = target.ToWordDocumentResult();
        using WordDocument sourceWord = sourceConversion.Value;
        using WordDocument targetWord = targetConversion.Value;
        WordComparisonResult comparison = WordDocumentComparer.CompareStructure(
            sourceWord,
            targetWord,
            options,
            sourceLabel: "source.rtf",
            targetLabel: "target.rtf");
        var report = new RtfConversionReport();
        report.Merge(sourceConversion.Report);
        report.Merge(targetConversion.Report);
        report.Add(RtfConversionSeverity.Information, "RtfWordCompareCompleted", "RTF documents were compared by OfficeIMO.Word.", RtfConversionAction.Preserved, feature: "compare");
        return new RtfWordWorkflowResult<WordComparisonResult>(source, comparison, report);
    }

    private static RtfWordWorkflowResult<T> Complete<T>(
        WordDocument word,
        T workflowResult,
        RtfConversionReport inputReport,
        string code,
        string message) {
        RtfConversionResult<RtfDocument> output = word.ToRtfDocumentResult();
        var report = new RtfConversionReport();
        report.Merge(inputReport);
        report.Merge(output.Report);
        report.Add(RtfConversionSeverity.Information, code, message, RtfConversionAction.Preserved, feature: "word-workflow");
        return new RtfWordWorkflowResult<T>(output.Value, workflowResult, report);
    }

}
