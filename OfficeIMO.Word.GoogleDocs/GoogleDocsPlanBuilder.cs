using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    internal static class GoogleDocsPlanBuilder {
        internal static GoogleDocsTranslationPlan Build(WordDocument document, GoogleDocsSaveOptions options) {
            var report = new TranslationReport();
            var plan = new GoogleDocsTranslationPlan(report) {
                SectionCount = document.Sections.Count,
                ParagraphCount = document.Paragraphs.Count,
                TableCount = document.Tables.Count,
                NestedTableCount = Math.Max(0, document.TablesIncludingNestedTables.Count - document.Tables.Count),
                ListCount = document.Lists.Count,
                ImageCount = document.Images.Count,
                ChartCount = document.Charts.Count,
                CommentCount = document.Comments.Count,
                FootnoteCount = document.FootNotes.Count,
                WatermarkCount = document.Watermarks.Count,
                SmartArtCount = document.SmartArts.Count,
                ShapeCount = document.Shapes.Count,
                TextBoxCount = document.TextBoxes.Count,
                EmbeddedObjectCount = document.EmbeddedObjects.Count,
                StructuredDocumentTagCount = document.StructuredDocumentTags.Count,
                CheckBoxCount = document.CheckBoxes.Count,
                DatePickerCount = document.DatePickers.Count,
                DropDownListCount = document.DropDownLists.Count,
                ComboBoxCount = document.ComboBoxes.Count,
                EquationCount = document.ParagraphsEquations.Count,
            };

            foreach (var section in document.Sections) {
                if (section.Header?.Default != null) plan.HeaderCount++;
                if (section.Header?.Even != null) plan.HeaderCount++;
                if (section.Header?.First != null) plan.HeaderCount++;
                if (section.Footer?.Default != null) plan.FooterCount++;
                if (section.Footer?.Even != null) plan.FooterCount++;
                if (section.Footer?.First != null) plan.FooterCount++;
            }

            plan.RequiresIndexAwareSections = plan.SectionCount > 1 || plan.HeaderCount > 0 || plan.FooterCount > 0 || plan.FootnoteCount > 0;

            if (plan.ShapeCount > 0 || plan.TextBoxCount > 0) {
                var message = options.FlattenFloatingContent
                    ? "Floating shapes and text boxes are planned for flattening because Google Docs translation is inline-first."
                    : "Floating shapes and text boxes require a dedicated layout strategy before export can be implemented.";
                report.Add(TranslationSeverity.Warning, "FloatingLayout", message);
            }

            if (plan.ChartCount > 0) {
                var message = options.RasterizeWordCharts
                    ? "Word charts are planned for rasterization or alternative embedding because there is no direct Word-chart target in Google Docs."
                    : "Word charts require a dedicated export strategy before native Google Docs export can be implemented.";
                report.Add(TranslationSeverity.Warning, "Charts", message);
            }

            if (plan.SmartArtCount > 0) {
                report.Add(TranslationSeverity.Warning, "SmartArt", "SmartArt has no direct Google Docs equivalent and is expected to be flattened or rendered.");
            }

            if (plan.StructuredDocumentTagCount > 0 || plan.CheckBoxCount > 0 || plan.DatePickerCount > 0 || plan.DropDownListCount > 0 || plan.ComboBoxCount > 0) {
                report.Add(TranslationSeverity.Warning, "ContentControls", "Word content controls do not map cleanly to Google Docs and will need flattening or alternate representations.");
            }

            if (plan.EmbeddedObjectCount > 0) {
                report.Add(TranslationSeverity.Warning, "EmbeddedObjects", "Embedded OLE objects do not have a direct Google Docs representation.");
            }

            if (plan.WatermarkCount > 0) {
                report.Add(TranslationSeverity.Info, "Watermarks", "Watermarks likely need to be ignored, flattened, or restated in headers.");
            }

            if (plan.CommentCount > 0 && options.PreserveCommentsViaDriveApi) {
                report.Add(TranslationSeverity.Info, "Comments", "Comments are expected to use the Drive comments API, which does not preserve rich Google Docs anchoring semantics.");
            } else if (plan.CommentCount > 0) {
                report.Add(TranslationSeverity.Warning, "Comments", "Comments exist in the document but comment export is not enabled in the current plan.");
            }

            if (plan.EquationCount > 0) {
                report.Add(TranslationSeverity.Warning, "Equations", "Equation writing needs a verified Google Docs request strategy and should be treated as deferred.");
            }

            if (plan.NestedTableCount > 0) {
                report.Add(TranslationSeverity.Info, "NestedTables", "Nested tables are present and may require staged insertion logic in Google Docs.");
            }

            return plan;
        }
    }
}
