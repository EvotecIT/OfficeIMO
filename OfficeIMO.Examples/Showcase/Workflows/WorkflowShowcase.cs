using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Runs the website's document walkthroughs and validates their editable outputs.</summary>
internal static class WorkflowShowcase {
    private sealed record Example(string Format, Action<string> Generate);

    private static readonly Dictionary<string, Example> Examples = new(StringComparer.Ordinal) {
        ["word-project-charter"] = new("word", ProjectCharter.Create),
        ["word-meeting-minutes"] = new("word", MeetingMinutes.Create),
        ["word-onboarding-checklist"] = new("word", OnboardingChecklist.Create),
        ["word-policy-brief"] = new("word", PolicyBrief.Create),
        ["word-decision-record"] = new("word", WordDecisionRecord.Create),
        ["excel-budget-scenarios"] = new("excel", BudgetScenarios.Create),
        ["excel-named-range-pricing"] = new("excel", NamedRangePricing.Create),
        ["excel-project-workbook"] = new("excel", ProjectWorkbook.Create),
        ["excel-expense-register"] = new("excel", ExpenseRegister.Create),
        ["excel-object-export"] = new("excel", ObjectExport.Create),
        ["pdf-dispatch-note"] = new("pdf", DispatchNote.Create),
        ["pdf-workshop-pack"] = new("pdf", WorkshopPack.Create),
        ["pdf-service-catalog"] = new("pdf", ServiceCatalog.Create),
        ["pdf-weekly-brief"] = new("pdf", WeeklyBrief.Create),
        ["powerpoint-project-roadmap"] = new("powerpoint", ProjectRoadmap.Create),
        ["powerpoint-training-quiz"] = new("powerpoint", TrainingQuiz.Create),
        ["powerpoint-service-overview"] = new("powerpoint", ServiceOverview.Create),
        ["markdown-incident-runbook"] = new("markdown", IncidentRunbook.Create),
        ["markdown-decision-record"] = new("markdown", MarkdownDecisionRecord.Create),
        ["markdown-api-handover"] = new("markdown", ApiHandover.Create),
        ["word-research-brief"] = new("word", ResearchBrief.Create),
        ["word-service-intake"] = new("word", ServiceIntake.Create),
        ["word-mixed-orientation-report"] = new("word", MixedOrientationReport.Create),
        ["excel-inventory-reorder"] = new("excel", InventoryReorder.Create),
        ["excel-training-matrix"] = new("excel", TrainingMatrix.Create),
        ["excel-operational-trends"] = new("excel", OperationalTrends.Create),
        ["pdf-equipment-inspection"] = new("pdf", EquipmentInspection.Create),
        ["pdf-linked-resource-guide"] = new("pdf", LinkedResourceGuide.Create),
        ["pdf-navigable-handbook"] = new("pdf", NavigableHandbook.Create),
        ["powerpoint-results-briefing"] = new("powerpoint", ResultsBriefing.Create),
        ["powerpoint-branching-playbook"] = new("powerpoint", BranchingPlaybook.Create),
        ["markdown-migration-guide"] = new("markdown", MigrationGuide.Create),
        ["markdown-data-dictionary"] = new("markdown", DataDictionary.Create),
        ["html-course-certificate"] = new("html", CourseCertificate.Create),
        ["html-service-level-report"] = new("html", ServiceLevelReport.Create),
        ["opendocument-service-proposal"] = new("opendocument", OpenProposal.Create),
        ["opendocument-workshop-agenda"] = new("opendocument", OpenWorkshopAgenda.Create),
        ["rtf-support-bulletin"] = new("rtf", SupportBulletin.Create),
        ["visio-approval-swimlane"] = new("visio", ApprovalSwimlane.Create),
        ["visio-team-organisation"] = new("visio", TeamOrganisation.Create)
    };

    internal static void Run(string documentsRoot, string? exampleId = null) {
        if (exampleId is not null && !Examples.ContainsKey(exampleId)) {
            throw new ArgumentException("Unknown showcase example: " + exampleId, nameof(exampleId));
        }
        foreach (var entry in Examples) {
            if (exampleId is not null && entry.Key != exampleId) continue;
            string folder = Path.Combine(documentsRoot, "Workflows", entry.Key);
            Directory.CreateDirectory(folder);
            entry.Value.Generate(folder);
            ValidateAndPreview(entry.Value.Format, folder);
            Console.WriteLine("Showcase workflow complete: " + entry.Key);
        }
    }

    private static void ValidateAndPreview(string format, string folder) {
        string preview = Path.Combine(folder, "preview.pdf");
        switch (format) {
            case "word":
                using (WordDocument document = WordDocument.Load(Path.Combine(folder, "example.docx"))) {
                    var errors = document.ValidateDocument();
                    if (errors.Count > 0) throw new InvalidOperationException(errors[0].Description);
                    document.SaveAsPdf(preview);
                }
                break;
            case "excel":
                using (ExcelDocument document = ExcelDocument.Load(Path.Combine(folder, "example.xlsx"))) {
                    if (!document.DocumentIsValid) throw new InvalidOperationException("Invalid showcase workbook: " + folder);
                    document.InspectFormulas().EnsureAllSupported().EnsureAllHaveCachedResults().EnsureNoDependencyIssues();
                }
                break;
            case "powerpoint":
                using (PowerPointPresentation presentation = PowerPointPresentation.Load(Path.Combine(folder, "example.pptx"))) {
                    var errors = presentation.ValidateDocument();
                    if (errors.Count > 0) throw new InvalidOperationException(errors[0].Description);
                    presentation.SaveAsPdf(preview);
                    presentation.Slides[0].ExportImage(OfficeImageExportFormat.Png)
                        .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
                    for (int index = 1; index < presentation.Slides.Count; index++) {
                        presentation.Slides[index].ExportImage(OfficeImageExportFormat.Png)
                            .Save(Path.Combine(folder, $"slide-{index + 1}.png"), OfficeImageExportFileConflictPolicy.Replace);
                    }
                }
                break;
            case "markdown":
                MarkdownDoc.Load(Path.Combine(folder, "example.md")).SaveAsPdf(preview, new MarkdownToPdfOptions {
                    Theme = MarkdownVisualTheme.Report().WithColorScheme(MarkdownColorSchemeKind.Blue)
                });
                break;
            case "opendocument":
                OdtDocument.Load(Path.Combine(folder, "example.odt")).ToPdfDocumentResult().Save(preview);
                break;
            case "rtf":
                RtfDocument.Load(Path.Combine(folder, "example.rtf")).ToPdfDocumentResult().Save(preview);
                break;
            case "visio":
                var diagram = VisioDocument.Load(Path.Combine(folder, "example.vsdx"));
                var issues = diagram.Validate();
                if (issues.Count > 0) throw new InvalidOperationException(issues[0]);
                diagram.ExportImage(OfficeImageExportFormat.Png)
                    .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
                diagram.ExportImage(OfficeImageExportFormat.Svg)
                    .Save(Path.Combine(folder, "preview.svg"), OfficeImageExportFileConflictPolicy.Replace);
                break;
        }
    }
}
