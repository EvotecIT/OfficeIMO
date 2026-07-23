using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.Examples {
    internal static class Program {
        private static void Setup(string path) {
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            }
        }

        private static bool HasArgument(string[] args, string value) {
            return Array.Exists(args, arg => string.Equals(arg, value, StringComparison.OrdinalIgnoreCase));
        }

        private static bool HasOption(string[] args, string name) {
            string prefix = name + "=";
            return Array.Exists(args, arg =>
                string.Equals(arg, name, StringComparison.OrdinalIgnoreCase) ||
                arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }

        private static string? GetArgumentValue(string[] args, string name) {
            for (int i = 0; i < args.Length; i++) {
                if (!string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal)) {
                    return args[i + 1];
                }

                return null;
            }

            string prefix = name + "=";
            string? combined = args.FirstOrDefault(arg => arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
            return combined == null ? null : combined.Substring(prefix.Length);
        }

        private static void RunPdfExamples(string folderPath) {
            Pdf.BasicPdf.Example_Pdf_HelloWorld(folderPath, false);
            Pdf.WriterDefaults.Example_Pdf_DefaultStyles(folderPath, false);
            Pdf.WriterStyledRuns.Example_Pdf_StyledRuns(folderPath, false);
            Pdf.WriterListsTables.Example_Pdf_BulletsAndTable(folderPath, false);
            Pdf.WriterHeadersFooters.Example_Pdf_PageNumbers(folderPath, false);
            Pdf.LinksAndRules.Example_Pdf_LinksAndRules(folderPath, false);
            Pdf.WriterStyleCheatsheet.Example_Pdf_StyleCheatsheet(folderPath, false);
            Pdf.DrawingGalleryPdf.Example_Pdf_DrawingGallery(folderPath, false);
            Pdf.RowColumnsPdf.Example_Pdf_RowColumns(folderPath, false);
            Pdf.TableStyleGalleryPdf.Example_Pdf_TableStyleGallery(folderPath, false);
            Pdf.ProfessionalReportPdf.Example_Pdf_ProfessionalReport(folderPath, false);
            Pdf.ShowcaseStatementPdf.Example_Pdf_ShowcaseStatement(folderPath, false);
            Pdf.ShowcaseDashboardPdf.Example_Pdf_ShowcaseDashboard(folderPath, false);
            Pdf.ShowcaseManipulationPdf.Example_Pdf_ShowcaseManipulation(folderPath, false);
        }

        private static void RunPdfShowcaseExamples(string folderPath) {
            Pdf.ShowcaseStatementPdf.Example_Pdf_ShowcaseStatement(folderPath, false);
            Pdf.ShowcaseDashboardPdf.Example_Pdf_ShowcaseDashboard(folderPath, false);
            Pdf.ShowcaseManipulationPdf.Example_Pdf_ShowcaseManipulation(folderPath, false);
        }

        private static void RunVisioShowcaseExamples(string folderPath, string[] args) {
            if (HasArgument(args, "--visio-export")
                || HasArgument(args, "--visio-showcase-export")
                || HasArgument(args, "--visio-preview")) {
                throw new InvalidOperationException(
                    "The Visio desktop preview switches were removed from OfficeIMO.Examples. " +
                    "Use --visio-native-preview for dependency-free SVG/PNG previews, or run the Visio desktop baseline tests/manual workflow when Microsoft Visio comparison artifacts are needed.");
            }

            bool exportNativePreviews = HasArgument(args, "--visio-native-preview")
                || HasArgument(args, "--visio-native-export");
            bool openVisio = HasArgument(args, "--open-visio")
                || HasArgument(args, "--visio-open");

            Visio.VisioShowcase.Example_VisioShowcase(folderPath, openVisio, exportNativePreviews);
        }

        private static void RunPowerPointExamples(string folderPath) {
            DateTime startedUtc = DateTime.UtcNow.AddSeconds(-2);

            PowerPoint.BasicPowerPointDocument.Example_BasicPowerPoint(folderPath, false);
            PowerPoint.AdvancedPowerPoint.Example_AdvancedPowerPoint(folderPath, false);
            PowerPoint.ModernPowerPointDeck.Example_ModernPowerPointDeck(folderPath, false);
            PowerPoint.DesignerPowerPointDeck.Example_DesignerPowerPointDeck(folderPath, false);
            PowerPoint.DesignBriefRecommendationsPowerPoint.Example_DesignBriefRecommendationsPowerPoint(folderPath, false);
            PowerPoint.DeckPlanAdvisorPowerPoint.Example_DeckPlanAdvisorPowerPoint(folderPath, false);
            PowerPoint.LayoutStrategyComparisonPowerPoint.Example_LayoutStrategyComparisonPowerPoint(folderPath, false);
            PowerPoint.DirectEditingPowerPoint.Example(folderPath, false);
            PowerPoint.ShapesPowerPoint.Example_PowerPointShapes(folderPath, false);
            PowerPoint.SlidesManagementPowerPoint.Example_SlidesManagement(folderPath, false);
            PowerPoint.SectionsWithoutRepairPowerPoint.Example_PowerPointSectionsWithoutRepair(folderPath, false);
            PowerPoint.TablesPowerPoint.Example_PowerPointTables(folderPath, false);
            PowerPoint.TextFormattingPowerPoint.Example_TextFormattingPowerPoint(folderPath, false);
            PowerPoint.ThemeAndLayoutPowerPoint.Example_PowerPointThemeAndLayout(folderPath, false);
            PowerPoint.TransitionsThemesPowerPoint.Example_TransitionsThemes(folderPath, false);
            PowerPoint.UpdatePicturePowerPoint.Example_PowerPointUpdatePicture(folderPath, false);
            PowerPoint.ValidateDocument.Example(folderPath, false);
            PowerPoint.TestLazyInit.Example_TestLazyInit(folderPath, false);
            PowerPoint.EndToEndPowerPointProof.Example_EndToEndPowerPointProof(folderPath, false);

            ValidateGeneratedPowerPointDecks(folderPath, startedUtc);
        }

        private static void ValidateGeneratedPowerPointDecks(string folderPath, DateTime startedUtc) {
            List<string> failures = new();
            List<string> generatedFiles = Directory
                .EnumerateFiles(folderPath, "*.pptx")
                .Where(file => File.GetLastWriteTimeUtc(file) >= startedUtc)
                .OrderBy(file => file, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (string filePath in generatedFiles) {
                using global::OfficeIMO.PowerPoint.PowerPointPresentation presentation =
                    global::OfficeIMO.PowerPoint.PowerPointPresentation.Load(filePath);
                List<DocumentFormat.OpenXml.Validation.ValidationErrorInfo> errors =
                    presentation.ValidateDocument().ToList();
                if (errors.Count == 0) {
                    continue;
                }

                string details = string.Join("; ", errors.Take(3).Select(error => error.Description));
                failures.Add($"{Path.GetFileName(filePath)}: {errors.Count} validation error(s). {details}");
            }

            if (failures.Count > 0) {
                throw new InvalidOperationException(
                    "One or more PowerPoint examples generated invalid Open XML." +
                    Environment.NewLine +
                    string.Join(Environment.NewLine, failures));
            }

            Console.WriteLine($"    Validation: {generatedFiles.Count} generated PowerPoint deck(s) passed Open XML validation.");
        }

        static void Main(string[] args) {
            string baseFolder = Path.TrimEndingDirectorySeparator(AppContext.BaseDirectory);
            Directory.SetCurrentDirectory(baseFolder);
            string templatesPath = Path.Combine(baseFolder, "Templates");
            string folderPath = Path.Combine(baseFolder, "Documents");
            Setup(folderPath);
            string? googleSupportMatrixPath = GetArgumentValue(args, "--google-support-matrix");
            if (!string.IsNullOrWhiteSpace(googleSupportMatrixPath)) {
                Google.GoogleWorkspaceSupportMatrixWriter.WriteTo(googleSupportMatrixPath!);
                Console.WriteLine($"Google Workspace support matrix written to {Path.GetFullPath(googleSupportMatrixPath!)}");
                return;
            }
            if (HasArgument(args, "--google-workspace")) {
                Google.GoogleDocsExamples.Example_Plan(folderPath);
                Google.GoogleSheetsExamples.Example_Plan(folderPath);
                Google.GoogleSlidesExamples.Example_Plan(folderPath);
                Google.GoogleDocsExamples.Example_ExportAsync(folderPath).GetAwaiter().GetResult();
                Google.GoogleSheetsExamples.Example_ExportAsync(folderPath).GetAwaiter().GetResult();
                Google.GoogleSlidesExamples.Example_ExportAsync(folderPath).GetAwaiter().GetResult();
                return;
            }
            if (HasArgument(args, "--onenote")) {
                OneNote.OfflineOneNoteExample.Example(folderPath);
                return;
            }

            if (HasArgument(args, "--opendocument")) {
                OpenDocument.OpenDocumentMilestones.Example(folderPath);
                return;
            }

            if (HasArgument(args, "--modern-powerpoint")) {
                PowerPoint.ModernPowerPointDeck.Example_ModernPowerPointDeck(folderPath, false);
                return;
            }

            if (HasArgument(args, "--designer-powerpoint")) {
                PowerPoint.DesignerPowerPointDeck.Example_DesignerPowerPointDeck(folderPath, false);
                return;
            }

            if (HasArgument(args, "--powerpoint-design-brief")) {
                PowerPoint.DesignBriefRecommendationsPowerPoint.Example_DesignBriefRecommendationsPowerPoint(folderPath, false);
                return;
            }

            if (HasArgument(args, "--powerpoint-deck-plan")) {
                PowerPoint.DeckPlanAdvisorPowerPoint.Example_DeckPlanAdvisorPowerPoint(folderPath, false);
                return;
            }

            if (HasArgument(args, "--powerpoint-layout-strategy")) {
                PowerPoint.LayoutStrategyComparisonPowerPoint.Example_LayoutStrategyComparisonPowerPoint(folderPath, false);
                return;
            }

            if (HasArgument(args, "--powerpoint-e2e")) {
                PowerPoint.EndToEndPowerPointProof.Example_EndToEndPowerPointProof(folderPath, false);
                return;
            }

            if (HasArgument(args, "--powerpoint")) {
                RunPowerPointExamples(folderPath);
                return;
            }

            if (HasArgument(args, "--excel-report-workflow")) {
                Excel.ReportWorkflow.Example(folderPath, false);
                return;
            }

            if (HasArgument(args, "--html-direct")) {
                Html.Html.Example_Html11_DirectOutputs(folderPath, HasArgument(args, "--open-pdf"));
                return;
            }

            if (HasArgument(args, "--pdf-professional")) {
                Pdf.ProfessionalReportPdf.Example_Pdf_ProfessionalReport(folderPath, false);
                return;
            }

            if (HasArgument(args, "--pdf-table-styles")) {
                Pdf.TableStyleGalleryPdf.Example_Pdf_TableStyleGallery(folderPath, false);
                return;
            }

            if (HasArgument(args, "--pdf-showcase")) {
                RunPdfShowcaseExamples(folderPath);
                return;
            }

            if (HasArgument(args, "--pdf")) {
                RunPdfExamples(folderPath);
                return;
            }

            if (HasArgument(args, "--markdown-advanced") || HasArgument(args, "--markdown-visual-fallback")) {
                Word.Converters.Markdown06_AdvancedWordRoundTrip.Example(folderPath, HasArgument(args, "--open-word"));
                return;
            }

            if (HasArgument(args, "--word-market-readiness")) {
                Word.MarketReadinessProofGallery.Example_GenerateWordMarketReadinessProof(folderPath, HasArgument(args, "--open-word"));
                return;
            }

            if (HasArgument(args, "--word-mail-merge-workflows")) {
                Word.MailMerge.Example_MailMergeWorkflowGallery(folderPath, HasArgument(args, "--open-word"));
                return;
            }

            if (HasArgument(args, "--word-review-reports")) {
                Word.ReviewReports.Example_ReviewReportWorkflow(folderPath, HasArgument(args, "--open-word"));
                return;
            }

            if (HasArgument(args, "--word-comparison-reports")) {
                Word.CompareDocuments.Example_ReportAndRedlineWorkflow(folderPath, HasArgument(args, "--open-word"));
                return;
            }

            if (HasArgument(args, "--word-signature-preflight")) {
                Word.SignaturePreflight.Example_SignaturePreflightWorkflow(folderPath, HasArgument(args, "--open-word"));
                return;
            }

            if (HasArgument(args, "--visio-premium") || HasArgument(args, "--premium-visio")) {
                Visio.PremiumVisioShowcase.Example_PremiumVisioShowcase(folderPath, HasArgument(args, "--open-visio") || HasArgument(args, "--visio-open"));
                return;
            }

            if (HasArgument(args, "--visio-showcase") || HasArgument(args, "--visio")) {
                RunVisioShowcaseExamples(folderPath, args);
                return;
            }

            if (HasArgument(args, "--visio-external-stencils")) {
                string? stencilPack = GetArgumentValue(args, "--visio-stencil-pack")
                    ?? Environment.GetEnvironmentVariable("OFFICEIMO_VISIO_STENCIL_PACK");
                if (string.IsNullOrWhiteSpace(stencilPack)) {
                    throw new InvalidOperationException("Provide a .vssx/.vsdx/.vstx path with --visio-stencil-pack <path> or OFFICEIMO_VISIO_STENCIL_PACK.");
                }

                Visio.ExternalStencilPack.Example_ExternalStencilPack(folderPath, HasArgument(args, "--open-visio") || HasArgument(args, "--visio-open"), stencilPack);
                return;
            }

            if (HasArgument(args, "--visio-installed-stencils")) {
                Visio.InstalledVisioStencils.Example_InstalledVisioStencils(folderPath, HasArgument(args, "--open-visio") || HasArgument(args, "--visio-open"));
                return;
            }

            if (HasOption(args, "--visio-integration-stencils") || HasOption(args, "--visio-microsoft-integration-stencils")) {
                string? stencilPackPath = Visio.MicrosoftIntegrationAzureStencils.ResolveConfiguredPackPath(args);
                if (string.IsNullOrWhiteSpace(stencilPackPath)) {
                    throw new InvalidOperationException("Provide a Microsoft Integration/Azure stencil pack file or root directory with --visio-integration-stencils <path> or OFFICEIMO_VISIO_INTEGRATION_STENCILS.");
                }

                Visio.MicrosoftIntegrationAzureStencils.Example_MicrosoftIntegrationAzureStencils(folderPath, HasArgument(args, "--open-visio") || HasArgument(args, "--visio-open"), stencilPackPath);
                return;
            }

            if (HasOption(args, "--visio-stencil-gallery") || HasOption(args, "--visio-stencil-gallery-pack")) {
                string? stencilGalleryPath = Visio.ExternalStencilGallery.ResolveConfiguredGalleryPath(args);
                if (string.IsNullOrWhiteSpace(stencilGalleryPath)) {
                    throw new InvalidOperationException("Provide a .vssx/.vsdx/.vstx file or root directory with --visio-stencil-gallery <path> or OFFICEIMO_VISIO_STENCIL_GALLERY.");
                }

                Visio.ExternalStencilGallery.Example_ExternalStencilGallery(folderPath, HasArgument(args, "--open-visio") || HasArgument(args, "--visio-open"), stencilGalleryPath);
                return;
            }

            if (HasArgument(args, "--visio-graph")) {
                Visio.GraphDiagramBuilder.Example_GraphDiagramBuilder(folderPath, HasArgument(args, "--open-visio") || HasArgument(args, "--visio-open"));
                return;
            }

            // Visio - Core Examples
            // Visio.BasicVisioDocument.Example_BasicVisio(folderPath, false);
            // Visio.FlowchartBuilder.Example_FlowchartBuilder(folderPath, false);
            // Visio.BlockDiagramBuilder.Example_BlockDiagramBuilder(folderPath, false);
            // Visio.DependencyDiagramBuilder.Example_DependencyDiagramBuilder(folderPath, false);
            // Visio.GraphDiagramBuilder.Example_GraphDiagramBuilder(folderPath, false);
            // Visio.MicrosoftIntegrationAzureStencils.Example_MicrosoftIntegrationAzureStencils(folderPath, false, @"C:\StencilPacks\Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio");
            // Visio.ExternalStencilGallery.Example_ExternalStencilGallery(folderPath, false, @"C:\StencilPacks\Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio");
            // Visio.ArchitectureDiagramBuilder.Example_ArchitectureDiagramBuilder(folderPath, false);
            // Visio.NetworkDiagramBuilder.Example_NetworkDiagramBuilder(folderPath, false);
            // Visio.NetworkTopologyDiagramBuilder.Example_NetworkTopologyDiagramBuilder(folderPath, false);
            // Visio.SwimlaneDiagramBuilder.Example_SwimlaneDiagramBuilder(folderPath, false);
            // Visio.OrgChartDiagramBuilder.Example_OrgChartDiagramBuilder(folderPath, false);
            // Visio.TimelineDiagramBuilder.Example_TimelineDiagramBuilder(folderPath, false);
            // Visio.SequenceDiagramBuilder.Example_SequenceDiagramBuilder(folderPath, false);
            // Visio.StencilCatalog.Example_StencilCatalog(folderPath, false);
            // Visio.QueryAndSelection.Example_QueryAndSelection(folderPath, false);
            // Visio.LayoutEditing.Example_LayoutEditing(folderPath, false);
            // Visio.LayerEditing.Example_LayerEditing(folderPath, false);
            // Visio.BackgroundPages.Example_BackgroundPages(folderPath, false);
            // Visio.PageSettings.Example_PageSettings(folderPath, false);
            // Visio.HyperlinkEditing.Example_HyperlinkEditing(folderPath, false);
            // Visio.ContainerEditing.Example_ContainerEditing(folderPath, false);
            // Visio.ShapeDataEditing.Example_ShapeDataEditing(folderPath, false);
            // Visio.ProtectionEditing.Example_ProtectionEditing(folderPath, false);
            // Visio.StyleThemes.Example_StyleThemes(folderPath, false);
            // Visio.ConnectorRouting.Example_ConnectorRouting(folderPath, false);
            // Visio.VisualQualityGallery.Example_VisualQualityGallery(folderPath, false);
            // Visio.ConnectRectangles.Example_ConnectRectangles(folderPath, false);
            // Visio.ConnectionPoints.Example_ConnectionPoints(folderPath, false);
            // Visio.ComprehensiveColoredShapes.Example_ComprehensiveColoredShapes(folderPath, false);
            // Visio.AllShapes.Example_AllShapes(folderPath, true);
            // Visio.AllShapesTyped.Example_AllShapes_Typed(folderPath, true);
            // Visio.AssetsCatalog.Example_ListAndExtractMasters(folderPath, false);
            // Visio.ReadVisioDocument.Example_ReadVisio(folderPath, false);
            // // Excel/BasicExcelFunctionality
            // Excel.BasicExcelFunctionality.BasicExcel_Example1(folderPath, false);
            // Excel.BasicExcelFunctionality.BasicExcel_Example2(folderPath, false);
            // Excel.BasicExcelFunctionality.BasicExcel_Example3(false);
            // // Excel/BasicExcelFunctionalityAsync
            // Excel.BasicExcelFunctionalityAsync.Example_ExcelAsync(folderPath).GetAwaiter().GetResult();
            // // Excel/AutoFit
            // Excel.AutoFit.Example(folderPath, false);
            // // Excel/AddTable
            // Excel.AddTable.Example(folderPath, false);
            // // Excel/AddTableMissingCells
            // Excel.AddTableMissingCells.Example(folderPath, false);
            // // Excel/AutoFilter
            // Excel.AutoFilter.Example(folderPath, false);
            // // Excel/Styles & Colors
            // Excel.StylesColors.Example(folderPath, false);
            // // Excel/Freeze
            // Excel.Freeze.Example(folderPath, false);
            // // Excel/ConditionalFormatting
            // Excel.ConditionalFormatting.Example(folderPath, false);
            // // Excel/ConcurrentWrites
            // Excel.ConcurrentWrites.Example(folderPath, false);
            // // Excel/ExcelConcurrentAccessAsync
            // Excel.ExcelConcurrentAccessAsync.Example_ExcelAsyncConcurrent(folderPath).GetAwaiter().GetResult();
            // // Excel/CellValues
            // Excel.CellValues.Example(folderPath, false);
            // // Excel/CellValuesParallel
            // Excel.CellValuesParallel.Example(folderPath, false);
            // // Excel/ValidateDocument
            // Excel.ValidateDocument.Example(folderPath, false);
            // // Excel/Fluent
            // Excel.FluentWorkbook.Example_FluentWorkbook(folderPath, false);
            // Excel.FluentWorkbook.Example_RangeBuilder(folderPath, false);
            // Excel.FluentWorkbook.Example_FluentWorkbook_AutoFilter(folderPath, false);
            // Excel.TestDateTime.Example_TestDateTime(folderPath, false);
            // // Excel/Read with custom converters
            // Excel.ReadWithConverters.Example(folderPath, false);
            // // Excel/Read presets and static helpers
            // Excel.ReadPresetsAndHelpers.Example(folderPath, false);
            // // Excel/Read for PowerShell consumption (emits JSON rows)
            // Excel.ReadForPowerShell.Example(folderPath, false);
            // // Excel/Feature preflight for read/edit/template/PDF workflow routing
            // Excel.FeaturePreflight.Example(folderPath, false);
            // // Excel/Report workflow with template, formulas, chart, pivot, preflight, and PDF
            // Excel.ReportWorkflow.Example(folderPath, false);
            // // Excel/PowerShell-style round trip: write → read → modify → write → JSON
            // Excel.PowerShellRoundTrip.Example(folderPath, false);
            // // Excel/Headers + Footers + Properties
            // Excel.HeadersFootersAndProperties.Example(folderPath, false);
            // Excel.DomainDetectiveReport.Example(folderPath, false);
            // // Excel: New Excelish Sheets demo (side-by-side comparison)
            // Excel.DomainDetectiveReportSheets.Example(folderPath, false);
            // // Excel: Classic baseline Sheets demo (explicit/standard techniques)
            // Excel.DomainDetectiveReportSheetsClassic.Example(folderPath, false);
            // // Google Workspace / Google Sheets
            // Google.GoogleSheetsExamples.Example_Plan(folderPath);
            // Google.GoogleSheetsExamples.Example_ExportAsync(folderPath).GetAwaiter().GetResult();
            // // Excel: Anchors and back-to-top demo
            // Excel.AnchorsAndBackToTop.Example(folderPath, false);
            // // Excel: Left-to-right multiple tables on same sheet
            // Excel.SheetComposerMultiTables.Example_LeftToRight(folderPath, true);
            // Excel.WrapText.Example(folderPath, false);
            // Excel.RowsFromObjects.Example(folderPath, false);
            // Excel.RowsFromObjectsPriorityProperties.Example(folderPath, false);
            // Excel.DomainDetectiveReportSheets.Example(folderPath, false);
            // // // Markdown: Anchors + Theme Toggle
            // Markdown.Markdown03_Anchors_Theme.Example_AnchorsAndTheme(folderPath, false);
            // // Markdown: TOC Layouts & Themes
            // Markdown.Markdown04_TocLayoutsAndThemes.Example_Toc_PanelTop(folderPath, false);
            // Markdown.Markdown04_TocLayoutsAndThemes.Example_Toc_SidebarLeft(folderPath, false);
            // Markdown.Markdown04_TocLayoutsAndThemes.Example_Toc_SidebarRight_ScrollSpy(folderPath, false);
            // Markdown.Markdown04_TocLayoutsAndThemes.Example_Toc_ScrollSpy_Long_IndigoTheme(folderPath, false);
            // // Markdown: Built-in HTML style gallery
            // Markdown.Markdown05_ThemesGallery.Example_Themes(folderPath, false);
            // // Markdown: One shared visual theme across Markdown, HTML, PDF, and Word
            // Markdown.Markdown10_VisualThemesAcrossFormats.Example_SharedVisualTheme(folderPath, false);
            // // Markdown: Custom parser/AST/HTML extensions
            // Markdown.Markdown07_Custom_Extensions.Example_Custom_Extensions(folderPath, false);
            // // Markdown: Delegate-based custom block parser extensions
            // Markdown.Markdown08_Custom_Block_Parsers.Example_Custom_Block_Parsers(folderPath, false);
            // // Markdown: Context-aware markdown write overrides
            // Markdown.Markdown09_Custom_Markdown_Write_Overrides.Example_Custom_Markdown_Write_Overrides(folderPath, false);
            // // Word ⇄ Markdown ⇄ HTML End-to-End
            // Word.EndToEnd.Word_EndToEnd.Example(folderPath, false);
            // // Markdown/DomainDetective report (mirrors the Excel structure)
            // Markdown.DomainDetectiveReportMarkdown.Example(folderPath, false);


            Excel.ChartsExcel.Charts_Basic(folderPath, true);
            Excel.ChartsExcel.Charts_ComboAndScatter(folderPath, true);



            // // PowerPoint
            PowerPoint.BasicPowerPointDocument.Example_BasicPowerPoint(folderPath, false);
            PowerPoint.AdvancedPowerPoint.Example_AdvancedPowerPoint(folderPath, false);
            PowerPoint.ModernPowerPointDeck.Example_ModernPowerPointDeck(folderPath, false);
            PowerPoint.DesignerPowerPointDeck.Example_DesignerPowerPointDeck(folderPath, false);
            PowerPoint.DesignBriefRecommendationsPowerPoint.Example_DesignBriefRecommendationsPowerPoint(folderPath, false);
            PowerPoint.DeckPlanAdvisorPowerPoint.Example_DeckPlanAdvisorPowerPoint(folderPath, false);
            PowerPoint.LayoutStrategyComparisonPowerPoint.Example_LayoutStrategyComparisonPowerPoint(folderPath, false);
            PowerPoint.DirectEditingPowerPoint.Example(folderPath, false);
            PowerPoint.ShapesPowerPoint.Example_PowerPointShapes(folderPath, false);
            PowerPoint.SlidesManagementPowerPoint.Example_SlidesManagement(folderPath, false);
            PowerPoint.SectionsWithoutRepairPowerPoint.Example_PowerPointSectionsWithoutRepair(folderPath, false);
            PowerPoint.TablesPowerPoint.Example_PowerPointTables(folderPath, false);
            PowerPoint.TextFormattingPowerPoint.Example_TextFormattingPowerPoint(folderPath, false);
            PowerPoint.ThemeAndLayoutPowerPoint.Example_PowerPointThemeAndLayout(folderPath, false);
            PowerPoint.TransitionsThemesPowerPoint.Example_TransitionsThemes(folderPath, false);
            PowerPoint.UpdatePicturePowerPoint.Example_PowerPointUpdatePicture(folderPath, false);
            PowerPoint.ValidateDocument.Example(folderPath, false);
            PowerPoint.TestLazyInit.Example_TestLazyInit(folderPath, false);
            // // Html/Html (consolidated set)
            // Html.Html.Example_Html01_LoadAndRoundTripBasics(folderPath, false);
            // Html.Html.Example_Html02_SaveAsHtmlFromWord(folderPath, false);
            // Html.Html.Example_Html03_TextFormatting(folderPath, false);
            // Html.Html.Example_Html04_ListsAndNumbering(folderPath, false);
            // Html.Html.Example_Html05_TablesComplex(folderPath, false);
            // Html.Html.Example_Html06_ImagesAllModes(folderPath, false);
            // Html.Html.Example_Html07_LinksAndAnchors(folderPath, false);
            // Html.Html.Example_Html08_SemanticsAndCitations(folderPath, false);
            // Html.Html.Example_Html09_CodePreWhitespace(folderPath, false);
            // Html.Html.Example_Html10_OptionsAndAsync(folderPath, false).GetAwaiter().GetResult();
            // Html.Html.Example_Html11_DirectOutputs(folderPath, false);
            // Html.Html.Example_Html00_AllInOne(folderPath, false);
            // // Markdown/Markdown
            // Markdown.Markdown.Example_MarkdownInterface(folderPath, false);
            // Markdown.Markdown.Example_MarkdownLists(folderPath, false);
            // Markdown.Markdown.Example_MarkdownRoundTrip(folderPath, false);
            // Markdown.Markdown.Example_MarkdownFootNotes(folderPath, false);
            // Markdown.Markdown.Example_MarkdownHeadingsBoldLinks(folderPath, false);
            // // Markdown/Builder & TOC
            // Markdown.Markdown01_Builder_Basics.Example_Builder_Readme(folderPath, false);
            // Markdown.Markdown01_Builder_Basics.Example_Scaffold_Readme(folderPath, false);
            // Markdown.Markdown02_DataToTableAndLists.Example_TablesAndLists(folderPath, false);
            // Markdown.Markdown02_DataToTableAndLists.Example_Toc(folderPath, false);
            // Markdown.Markdown02_DataToTableAndLists.Example_Table_FromAny_WithOptions(folderPath, false);
            // Markdown.Markdown02_DataToTableAndLists.Example_Table_FromSequence_WithSelectors(folderPath, false);
            // Markdown.Markdown02_DataToTableAndLists.Example_HeaderTransform_CustomAcronyms(folderPath, false);
            // Markdown.Markdown02_DataToTableAndLists.Example_Table_AutoAligners(folderPath, false);
            // Markdown.Markdown02_DataToTableAndLists.Example_TocForSection(folderPath, false);
            // // Word/AdvancedDocument
            // Word.AdvancedDocument.Example_AdvancedWord(folderPath, false);
            // Word.AdvancedDocument.Example_AdvancedWord2(folderPath, false);
            // Word/WebCompat quick galleries
            // Word.WebCompat.Example_TablesGallery(folderPath, false);
            // Word.WebCompat.Example_CoverTemplates_Basic(folderPath, false);
            // Word.WebCompat.Example_CoverWithConfidentialWatermark(folderPath, false);
            // Word/Background
            // Word.Background.Example_BackgroundImageAdvanced(folderPath, false);
            // Word.Background.Example_BackgroundImageSimple(folderPath, false);
            // // Word/BasicDocument
            // Word.BasicDocument.Example_BasicDocument(folderPath, false);
            // Word.BasicDocument.Example_BasicDocumentSaveAs1(folderPath, false);
            // Word.BasicDocument.Example_BasicDocumentSaveAs2(folderPath, false);
            // Word.BasicDocument.Example_BasicDocumentSaveAs3(folderPath, fa lse);
            // Word.BasicDocument.Example_BasicDocumentWithoutUsing(folderPath, false);
            // Word.BasicDocument.Example_BasicEmptyWord(folderPath, false);
            // Word.BasicDocument.Example_BasicLoadHamlet(templatesPath, folderPath, false);
            // Word.BasicDocument.Example_BasicWord(folderPath, false);
            // Word.BasicDocument.Example_BasicWord2(folderPath, false);
            // Word.BasicDocument.Example_BasicWordAsync(folderPath).GetAwaiter().GetResult();
            // Word.BasicDocument.Example_BasicWordWithBreaks(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithDefaultFontChange(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithDefaultStyleChange(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithLineSpacing(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithMargins(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithMarginsAndImage(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithMarginsInCentimeters(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithNewLines(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithPolishChars(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithSomeParagraphs(folderPath, false);
            // Word.BasicDocument.Example_BasicWordWithTabs(folderPath, false);
            // // Word/Bookmarks
            // Word.Bookmarks.Example_BasicWordWithBookmarks(folderPath, false);
            // // Word/BordersAndMargins
            // Word.BordersAndMargins.Example_BasicPageBorders1(folderPath, false);
            // Word.BordersAndMargins.Example_BasicPageBorders2(folderPath, false);
            // Word.BordersAndMargins.Example_BasicWordMarginsSizes(folderPath, false);
            // // Word/Charts
            // Word.Charts.Example_AddingMultipleCharts(folderPath, false);
            // Word.Charts.Example_Area3DChart(folderPath, false);
            // Word.Charts.Example_AreaChart(folderPath, false);
            // Word.Charts.Example_Bar3DChart(folderPath, false);
            // Word.Charts.Example_BarChart(folderPath, false);
            // Word.Charts.Example_ComboChart(folderPath, false);
            // Word.Charts.Example_Line3DChart(folderPath, false);
            // Word.Charts.Example_LineChart(folderPath, false);
            // Word.Charts.Example_Pie3DChart(folderPath, false);
            // Word.Charts.Example_PieChart(folderPath, false);
            // Word.Charts.Example_RadarChart(folderPath, false);
            // Word.Charts.Example_ScatterChart(folderPath, false);
            // Word.Charts.Example_Charts_PaletteAndSizing(folderPath, false);
            // // Word/CheckBoxes
            // Word.CheckBoxes.Example_BasicCheckBox(folderPath, false);
            // // Word/CitationsExamples
            // Word.CitationsExamples.Example_AdvancedCitations(folderPath, false);
            // Word.CitationsExamples.Example_BasicCitations(folderPath, false);
            // // Google Workspace / Google Docs
            // Google.GoogleDocsExamples.Example_Plan(folderPath);
            // Google.GoogleDocsExamples.Example_ExportAsync(folderPath).GetAwaiter().GetResult();
            // // Word/ComboBoxes
            // Word.ComboBoxes.Example_BasicComboBox(folderPath, false);
            // // Word/Comments
            // Word.Comments.Example_PlayingWithComments(folderPath, false);
            // Word.Comments.Example_RemoveCommentsAndTrack(folderPath, false);
            // Word.Comments.Example_ThreadedComments(folderPath, false);
            // // Word/CompareDocuments
            // Word.CompareDocuments.Example_BasicComparison(folderPath, false);
            // // Word/MergeDocuments
            // Word.MergeDocuments.Example_AppendDocument(folderPath, false);
            // // Word/ContentControls
            // Word.ContentControls.Example_AddContentControl(folderPath, false);
            // Word.ContentControls.Example_AdvancedContentControls(folderPath, false);
            // Word.ContentControls.Example_ContentControlsInTable(folderPath, false);
            // Word.ContentControls.Example_MultipleContentControls(folderPath, false);
            // Word.ContentControls.Example_FormattedContentControls(folderPath, false);
            // // Word/CleanupDocuments
            // Word.CleanupDocuments.CleanupDocuments_Sample01(false);
            // Word.CleanupDocuments.CleanupDocuments_Sample02(folderPath, false);
            // Word.CleanupDocuments.CleanupDocuments_Sample03(folderPath, false);
            // Word.CleanupDocuments.CleanupDocuments_Sample04(folderPath, false);
            // // Word/CoverPages
            // Word.CoverPages.Example_AddingCoverPage(folderPath, false);
            // Word.CoverPages.Example_AddingCoverPage2(folderPath, false);
            // // Word/CrossReferences
            // Word.CrossReferences.Example_BasicCrossReferences(folderPath, false);
            // // Word/CustomAndBuiltinProperties
            // Word.CustomAndBuiltinProperties.Example_BasicCustomProperties(folderPath, false);
            // Word.CustomAndBuiltinProperties.Example_BasicDocumentProperties(folderPath, false);
            // Word.CustomAndBuiltinProperties.Example_Load(false);
            // Word.CustomAndBuiltinProperties.Example_LoadDocumentWithProperties(false);
            // Word.CustomAndBuiltinProperties.Example_ReadWord(false);
            // Word.CustomAndBuiltinProperties.Example_ValidateDocument(folderPath);
            // Word.CustomAndBuiltinProperties.Example_ValidateDocument_BeforeSave();
            // // Word/DatePickers
            // Word.DatePickers.Example_BasicDatePicker(folderPath, false);
            // Word.DatePickers.Example_AdvancedDatePicker(folderPath, false);
            // // Word/DocumentVariablesExamples
            // Word.DocumentVariablesExamples.Example_AdvancedDocumentVariables(folderPath, false);
            // Word.DocumentVariablesExamples.Example_BasicDocumentVariables(folderPath, false);
            // // Word/DropDownLists
            // Word.DropDownLists.Example_BasicDropDownList(folderPath, false);
            // Word.DropDownLists.Example_AdvancedDropDownList(folderPath, false);
            // // Word/Embed
            // Word.Embed.Example_EmbedFileExcel(folderPath, templatesPath, false);
            // Word.Embed.Example_EmbedFileHTML(folderPath, templatesPath, false);
            // Word.Embed.Example_EmbedFileMultiple(folderPath, templatesPath, false);
            // Word.Embed.Example_EmbedFileRTF(folderPath, templatesPath, false);
            // Word.Embed.Example_EmbedFileRTFandHTML(folderPath, templatesPath, false);
            // Word.Embed.Example_EmbedFileRTFandHTMLandTOC(folderPath, templatesPath, false);
            // Word.Embed.Example_EmbedFragmentAfter(folderPath, false);
            // Word.Embed.Example_EmbedHTMLFragment(folderPath, false);
            // // Word/Equations
            // Word.Equations.Example_AddEquation(folderPath, false);
            // Word.Equations.Example_AddEquationExponent(folderPath, false);
            // Word.Equations.Example_AddEquationIntegral(folderPath, false);
            // // Word/Fields
            // Word.Fields.Example_CustomFormattedDateField(folderPath, false);
            // Word.Fields.Example_CustomFormattedHeaderDate(folderPath, false);
            // Word.Fields.Example_CustomFormattedTimeField(folderPath, false);
            // Word.Fields.Example_DocumentWithFields(folderPath, false);
            // Word.Fields.Example_DocumentWithFields02(folderPath, false);
            // Word.Fields.Example_FieldBuilderNested(folderPath, false);
            // //OfficeIMO.Examples.Word.Fields.Example_FieldBuilderSimple(folderPath, false);
            // Word.Fields.Example_FieldFormatAdvanced(folderPath, false);
            // Word.Fields.Example_FieldFormatRoman(folderPath, false);
            // Word.Fields.Example_FieldWithMultipleSwitches(folderPath, false);
            // // Word/FindAndReplace
            // Word.FindAndReplace.Example_FindAndReplace01(folderPath, false);
            // Word.FindAndReplace.Example_FindAndReplace02(folderPath, false);
            // Word.FindAndReplace.Example_FindAndReplace03(folderPath, false);
            // Word.FindAndReplace.Example_ReplaceTextWithHtmlFragment(folderPath, false);
            // // Word/Fluent
            // Word.FluentDocument.Example_FluentDocument(folderPath, false);
            // Word.FluentDocument.Example_FluentHeadersAndFooters(folderPath, false);
            // Word.FluentDocument.Example_FluentListBuilder(folderPath, false);
            // Word.FluentDocument.Example_FluentParagraphFormatting(folderPath, false);
            // Word.FluentDocument.Example_FluentReadHelpers(folderPath, false);
            // Word.FluentDocument.Example_FluentSectionLayout(folderPath, false);
            // Word.FluentDocument.Example_FluentTableBuilder(folderPath, false);
            // Word.FluentDocument.Example_FluentTextBuilder(folderPath, false);
            // // Word/Fonts
            // Word.Fonts.Example_EmbeddedAndBuiltinFonts(templatesPath, folderPath, false);
            // Word.Fonts.Example_EmbeddedFontStyle(templatesPath, folderPath, false);
            // Word.Fonts.Example_EmbedFont(templatesPath, folderPath, false);
            // Word.Fonts.Example_EmbedFontWithStyle(templatesPath, folderPath, false);
            // // Word/FootNotes
            // Word.FootNotes.Example_DocumentWithFootNotes(folderPath, false);
            // Word.FootNotes.Example_DocumentWithFootNotesEmpty(folderPath, false);
            // // Word/HeadersAndFooters
            // Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter(folderPath, false);
            // Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter0(folderPath, false);
            // Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter1(folderPath, false);
            // Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooterWithoutSections(folderPath, false);
            // Word.HeadersAndFooters.Sections1(folderPath, false);
            // // Word/HyperLinks
            // Word.HyperLinks.EasyExample(folderPath, false);
            // Word.HyperLinks.Example_AddingFields(folderPath, false);
            // Word.HyperLinks.Example_BasicWordWithHyperLinks(folderPath, false);
            // Word.HyperLinks.Example_BasicWordWithHyperLinksInTables(folderPath, false);
            // Word.HyperLinks.Example_FormattedHyperLinks(folderPath, false);
            // Word.HyperLinks.Example_FormattedHyperLinksAdvanced(folderPath, false);
            // Word.HyperLinks.Example_FormattedHyperLinksListReuse(folderPath, false);
            // // Word/Images
            // Word.Images.Example_AddingFixedImages(folderPath, false);
            // Word.Images.Example_AddingImages(folderPath, false);
            // Word.Images.Example_AddingImagesHeadersFooters(folderPath, false);
            // Word.Images.Example_AddingImagesInline(folderPath, false);
            // Word.Images.Example_AddingImagesMultipleTypes(folderPath, false);
            // Word.Images.Example_AddingImagesSample4(folderPath, false);
            // Word.Images.Example_AddingImagesSampleToTable(folderPath, false);
            // Word.Images.Example_ImageCroppingAdvanced(folderPath, false);
            // Word.Images.Example_ImageCroppingBasic(folderPath, false);
            // Word.Images.Example_ImageNewFeatures(folderPath, false);
            // Word.Images.Example_ImageTransparencyAdvanced(folderPath, false);
            // Word.Images.Example_ImageTransparencySimple(folderPath, false);
            // Word.Images.Example_ReadWordWithImages();
            // Word.Images.Example_ReadWordWithImagesAndDiffWraps();
            // // Word/Lists
            // Word.Lists.Example_BasicLists(folderPath, false);
            // Word.Lists.Example_BasicLists10(folderPath, false);
            // Word.Lists.Example_BasicLists11(folderPath, false);
            // Word.Lists.Example_BasicLists12(folderPath, false);
            // Word.Lists.Example_BasicLists2(folderPath, false);
            // Word.Lists.Example_BasicLists2Load(folderPath, false);
            // Word.Lists.Example_BasicLists3(folderPath, false);
            // Word.Lists.Example_BasicLists4(folderPath, false);
            // Word.Lists.Example_BasicLists6(folderPath, false);
            // Word.Lists.Example_BasicLists7(folderPath, false);
            // Word.Lists.Example_BasicLists8(folderPath, false);
            // Word.Lists.Example_BasicLists9(folderPath, false);
            // Word.Lists.Example_BasicListsWithChangedStyling(folderPath, false);
            // Word.Lists.Example_CloneList(folderPath, false);
            // Word.Lists.Example_CustomBulletColor(folderPath, false);
            // Word.Lists.Example_CustomList1(folderPath, false);
            // Word.Lists.Example_DetectListStyles(folderPath, false);
            // Word.Lists.Example_ListStartNumber(folderPath, false);
            // Word.Lists.Example_PictureBulletList(folderPath, false);
            // Word.Lists.Example_PictureBulletListAdvanced(folderPath, false);
            // // Word/LoadDocuments
            // Word.LoadDocuments.LoadWordDocument_Sample1(false);
            // Word.LoadDocuments.LoadWordDocument_Sample2(false);
            // Word.LoadDocuments.LoadWordDocument_Sample3(false);
            // // Word/Macros
            // Word.Macros.Example_AddMacroToExistingDocx(templatesPath, folderPath, false);
            // Word.Macros.Example_CreateDocmWithMacro(templatesPath, folderPath, false);
            // Word.Macros.Example_ExtractAndRemoveMacro(templatesPath, folderPath, false);
            // Word.Macros.Example_ListAndRemoveMacro(templatesPath, folderPath, false);
            // Word.Macros.Example_ListMacros(templatesPath, folderPath, false);
            // // Word/MailMerge
            // Word.MailMerge.Example_MailMergeAdvanced(folderPath, false);
            // Word.MailMerge.Example_MailMergeWorkflowGallery(folderPath, false);
            // Word.MailMerge.Example_MailMergeSimple(folderPath, false);
            // Word.MarketReadinessProofGallery.Example_GenerateWordMarketReadinessProof(folderPath, false);
            // // Word/PageBreaks
            // Word.PageBreaks.Example_PageBreaks(folderPath, false);
            // Word.PageBreaks.Example_PageBreaks1(folderPath, false);
            // // Word/PageNumbers
            // Word.PageNumbers.Example_PageNumbers1(folderPath, false);
            // Word.PageNumbers.Example_PageNumbers2(folderPath, false);
            // Word.PageNumbers.Example_PageNumbers3(folderPath, false);
            // Word.PageNumbers.Example_PageNumbers4(folderPath, false);
            // Word.PageNumbers.Example_PageNumbers5(folderPath, false);
            // Word.PageNumbers.Example_PageNumbers6(folderPath, false);
            // Word.PageNumbers.Example_PageNumbers7(folderPath, false);
            // Word.PageNumbers.Example_PageNumbers8(folderPath, false);
            // // Word/PageSettings
            // Word.PageSettings.Example_BasicSettings(folderPath, false);
            // Word.PageSettings.Example_PageOrientation(folderPath, false);
            // // Word/Paragraphs
            // Word.Paragraphs.Example_BasicParagraphs(folderPath, false);
            // Word.Paragraphs.Example_AddFormattedText(folderPath, false);
            // Word.Paragraphs.Example_BasicTabStops(folderPath, false);
            // Word.Paragraphs.Example_InsertParagraphAt(folderPath, false);
            // Word.Paragraphs.Example_BasicParagraphStyles(folderPath, false);
            // Word.Paragraphs.Example_InlineRunHelper(folderPath, false);
            // Word.Paragraphs.Example_MultipleParagraphsViaDifferentWays(folderPath, false);
            // Word.Paragraphs.Example_RunCharacterStylesSimple(folderPath, false);
            // Word.Paragraphs.Example_RunCharacterStylesAdvanced(folderPath, false);
            // Word.Paragraphs.Example_RegisterCustomParagraphStyle(folderPath, false);
            // Word.Paragraphs.Example_MultipleCustomParagraphStyles(folderPath, false);
            // Word.Paragraphs.Example_OverrideBuiltInParagraphStyle(folderPath, false);
            // Word.Paragraphs.Example_Word_Fluent_Paragraph_TextAndFormatting(folderPath, false);
            // // Word/Pdf
            // Word.Pdf.Example_HeaderFooterImages(folderPath, false);
            // Word.Pdf.Example_PdfInterface(folderPath, false);
            // Word.Pdf.Example_SaveAsPdf(folderPath, false);
            // Word.Pdf.Example_SaveAsPdfInMemory(folderPath, false);
            // Word.Pdf.Example_SaveAsPdfRelative(folderPath, false);
            // Word.Pdf.Example_SaveAsPdfWithHyperlinks(folderPath, false);
            // Word.Pdf.Example_SaveAsPdfWithMetadata(folderPath, false);
            // Word.Pdf.Example_SaveAsPdfWithFirstPartyOptions(folderPath, false);
            // Word.Pdf.Example_SaveLists(folderPath, false);
            // Word.Pdf.Example_TableStyles(folderPath, false);
            // Word.Pdf.Example_PdfCustomFonts(folderPath, false);
            // // Word/PictureControls
            // Word.PictureControls.Example_BasicPictureControl(folderPath, false);
            // // Word/Protection
            // Word.Protect.Example_FinalDocument(folderPath, false);
            // Word.Protect.Example_ReadOnlyEnforced(folderPath, false);
            // Word.Protect.Example_ReadOnlyRecommended(folderPath, false);
            // // Word/RepeatingSections
            // Word.RepeatingSections.Example_BasicRepeatingSection(folderPath, false);
            // // Word/Revisions
            // Word.Revisions.Example_ConvertRevisionsToMarkup(folderPath, false);
            // Word.Revisions.Example_TrackChangesToggle(folderPath, false);
            // Word.Revisions.Example_TrackedChanges(folderPath, false);
            // // Word/SaveToStream
            // Word.SaveToStream.Example_CreateInProvidedStream(folderPath, false);
            // Word.SaveToStream.Example_CreateInProvidedStreamAdvanced(folderPath, false);
            // Word.SaveToStream.Example_ToBytes(folderPath, false);
            // Word.SaveToStream.Example_ToStream(folderPath, false);
            // Word.SaveToStream.Example_SaveAsStream(folderPath, false);
            // Word.SaveToStream.Example_SaveToOriginalStream(folderPath, false);
            // Word.SaveToStream.Example_StreamDocumentProperties(folderPath, false);
            // // Word/Sections
            // Word.Sections.Example_BasicSections(folderPath, false);
            // Word.Sections.Example_BasicSections2(folderPath, false);
            // Word.Sections.Example_BasicSections3WithColumns(folderPath, false);
            // Word.Sections.Example_BasicWordWithSections(folderPath, false);
            // Word.Sections.Example_SectionsWithHeaders(folderPath, false);
            // Word.Sections.Example_SectionsWithHeadersDefault(folderPath, false);
            // Word.Sections.Example_SectionsWithParagraphs(folderPath, false);
            // Word.Sections.Example_CloneSection(folderPath, false);
            // // Word/Shapes
            // Word.Shapes.Example_AddBasicShape(folderPath, false);
            // Word.Shapes.Example_AddEllipseAndPolygon(folderPath, false);
            // Word.Shapes.Example_AddLine(folderPath, false);
            // Word.Shapes.Example_AddMultipleShapes(folderPath, false);
            // Word.Shapes.Example_LoadShapes(folderPath, false);
            // Word.Shapes.Example_RemoveShape(folderPath, false);
            // // Word/SmartArt
            // Word.SmartArt.Example_AddAdvancedSmartArt(folderPath, true);
            // Word.SmartArt.Example_AddBasicSmartArt(folderPath, true);
            // // Additional SmartArt examples from FixSmartArtShapes branch
            // Word.SmartArt.Example_AddCustomSmartArt1(folderPath, true);
            // Word.SmartArt.Example_AddCustomSmartArt2(folderPath, true);
            // // SmartArt edit flows
            // Word.SmartArt.Example_EditCustomSmartArt1(folderPath, true);
            // Word.SmartArt.Example_EditCustomSmartArt2(folderPath, true);
            // Word.SmartArt.Example_FlexibleBasicSmartArt_FullFlow(folderPath, true);
            // Word.SmartArt.Example_FlexibleCycleSmartArt_FullFlow(folderPath, true);
            // // Word/Tables
            // Word.Tables.Example_AllTables(folderPath, false);
            // Word.Tables.Example_BasicTables1(folderPath, false);
            // Word.Tables.Example_BasicTables10_StylesModificationWithCentimeters(folderPath, false);
            // Word.Tables.Example_BasicTables6(folderPath, false);
            // Word.Tables.Example_BasicTables8(folderPath, false);
            // Word.Tables.Example_BasicTables8_StylesModification(folderPath, false);
            // Word.Tables.Example_BasicTablesLoad1(folderPath, false);
            // Word.Tables.Example_BasicTablesLoad2(templatesPath, folderPath, false);
            // Word.Tables.Example_BasicTablesLoad3(templatesPath, false);
            // Word.Tables.Example_CloneTable(folderPath, false);
            // Word.Tables.Example_ConditionalFormattingAdvanced(folderPath, false);
            // Word.Tables.Example_ConditionalFormattingValues(folderPath, false);
            // Word.Tables.Example_DifferentTableSizes(folderPath, false);
            // Word.Tables.Example_InsertTableAfterAdvanced(folderPath, false);
            // Word.Tables.Example_InsertTableAfterSimple(folderPath, false);
            // Word.Tables.Example_InsertTableAfterWithXml(folderPath, false);
            // Word.Tables.Example_NestedTables(folderPath, false);
            // Word.Tables.Example_SplitHorizontally(folderPath, false);
            // Word.Tables.Example_SplitVertically(folderPath, false);
            // Word.Tables.Example_TableBorders(folderPath, false);
            // Word.Tables.Example_TableCellOptions(folderPath, false);
            // Word.Tables.Example_Tables(folderPath, false);
            // Word.Tables.Example_Tables1CopyRow(folderPath, false);
            // Word.Tables.Example_TablesAddedAfterParagraph(folderPath, false);
            // Word.Tables.Example_TablesWidthAndAlignment(folderPath, false);
            // Word.Tables.Example_UnifiedTableBorders(folderPath, false);
            // // Word/TOC
            // Word.TOC.Example_BasicTOC1(folderPath, false);
            // Word.TOC.Example_BasicTOC2(folderPath, false);
            // Word.TOC.Example_RemoveRegenerateTOC(folderPath, false);
            // // Word/UpdateFieldsSample
            // Word.UpdateFieldsSample.Example_UpdateFields(folderPath, false);
            // // Word/Watermark
            // Word.Watermark.Watermark_Remove(folderPath, false);
            // Word.Watermark.Watermark_Sample1(folderPath, false);
            // Word.Watermark.Watermark_Sample2(folderPath, false);
            // Word.Watermark.Watermark_Sample3(folderPath, false);
            // Word.Watermark.Watermark_SampleImage1(folderPath, false);
            // // Word/WordTextBox
            // Word.WordTextBox.Example_AddingTextbox(folderPath, false);
            // Word.WordTextBox.Example_AddingTextbox2(folderPath, false);
            // Word.WordTextBox.Example_AddingTextbox3(folderPath, false);
            // Word.WordTextBox.Example_AddingTextbox4(folderPath, false);
            // Word.WordTextBox.Example_AddingTextbox5(folderPath, false);
            // Word.WordTextBox.Example_AddingTextboxCentimeters(folderPath, false);
            // Word.WordTextBox.Example_TextBoxAutoFitOptions(folderPath, false);
            // // Word/XmlSerialization
            // Word.XmlSerialization.Example_XmlSerializationAdvanced(folderPath, false);
            // Word.XmlSerialization.Example_XmlSerializationBasic(folderPath, false);
        }
    }
}
