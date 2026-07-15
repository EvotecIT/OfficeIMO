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
                AddUnsupported(
                    report,
                    "FloatingLayout",
                    "DOCS.FLOATING_CONTENT",
                    plan.ShapeCount + plan.TextBoxCount,
                    options.UnsupportedFeatures.FloatingContent,
                    "floating shapes and text boxes",
                    options.InlineImageMode == GoogleDocsInlineImageMode.TemporaryPublicDriveLease);
            }

            if (plan.ChartCount > 0) {
                AddUnsupported(report, "Charts", "DOCS.CHART", plan.ChartCount, options.UnsupportedFeatures.Charts, "Word charts", options.InlineImageMode == GoogleDocsInlineImageMode.TemporaryPublicDriveLease);
            }

            if (plan.SmartArtCount > 0) {
                AddUnsupported(report, "SmartArt", "DOCS.SMART_ART", plan.SmartArtCount, options.UnsupportedFeatures.SmartArt, "SmartArt objects", options.InlineImageMode == GoogleDocsInlineImageMode.TemporaryPublicDriveLease);
            }

            if (plan.StructuredDocumentTagCount > 0 || plan.CheckBoxCount > 0 || plan.DatePickerCount > 0 || plan.DropDownListCount > 0 || plan.ComboBoxCount > 0) {
                int contentControlCount = plan.StructuredDocumentTagCount + plan.CheckBoxCount + plan.DatePickerCount + plan.DropDownListCount + plan.ComboBoxCount;
                AddUnsupported(report, "ContentControls", "DOCS.CONTENT_CONTROL", contentControlCount, options.UnsupportedFeatures.ContentControls, "Word content controls", options.InlineImageMode == GoogleDocsInlineImageMode.TemporaryPublicDriveLease);
            }

            if (plan.EmbeddedObjectCount > 0) {
                AddUnsupported(report, "EmbeddedObjects", "DOCS.EMBEDDED_OBJECT", plan.EmbeddedObjectCount, options.UnsupportedFeatures.EmbeddedObjects, "embedded OLE objects", options.InlineImageMode == GoogleDocsInlineImageMode.TemporaryPublicDriveLease);
            }

            if (plan.WatermarkCount > 0) {
                AddUnsupported(report, "Watermarks", "DOCS.WATERMARK", plan.WatermarkCount, options.UnsupportedFeatures.Watermarks, "watermarks", options.InlineImageMode == GoogleDocsInlineImageMode.TemporaryPublicDriveLease);
            }

            if (plan.CommentCount > 0) {
                if (options.Comments == GoogleDocsCommentMode.UnanchoredDriveComments) {
                    report.Add(
                        TranslationSeverity.Warning,
                        "Comments",
                        $"{plan.CommentCount} Word comment(s) will be created as unanchored Drive comments because Google editors do not honor Drive API anchors.",
                        code: "DOCS.COMMENT.UNANCHORED",
                        action: TranslationAction.Flatten,
                        count: plan.CommentCount);
                } else {
                    AddUnsupported(report, "Comments", "DOCS.COMMENT", plan.CommentCount, options.UnsupportedFeatures.Comments, "comments");
                }
            }

            if (plan.EquationCount > 0) {
                AddUnsupported(report, "Equations", "DOCS.EQUATION", plan.EquationCount, options.UnsupportedFeatures.Equations, "equations", options.InlineImageMode == GoogleDocsInlineImageMode.TemporaryPublicDriveLease);
            }

            if (plan.ImageCount > 0) {
                report.Add(
                    options.InlineImageMode == GoogleDocsInlineImageMode.Placeholder ? TranslationSeverity.Warning : TranslationSeverity.Info,
                    "InlineImages",
                    options.InlineImageMode == GoogleDocsInlineImageMode.Placeholder
                        ? "Inline images will remain readable placeholders because temporary public Drive staging was not explicitly enabled."
                        : "Inline images will use short-lived public Drive staging and the exporter will delete every staging file after Google Docs fetches it.",
                    code: options.InlineImageMode == GoogleDocsInlineImageMode.Placeholder
                        ? "DOCS.IMAGE.STAGING_DISABLED"
                        : "DOCS.IMAGE.TEMPORARY_PUBLIC_LEASE",
                    action: options.InlineImageMode == GoogleDocsInlineImageMode.Placeholder
                        ? TranslationAction.Skip
                        : TranslationAction.Preserve,
                    count: plan.ImageCount);
            }

            if (plan.NestedTableCount > 0) {
                report.Add(TranslationSeverity.Info, "NestedTables", "Nested tables are present and may require staged insertion logic in Google Docs.");
            }

            return plan;
        }

        private static void AddUnsupported(
            TranslationReport report,
            string feature,
            string codePrefix,
            int count,
            UnsupportedFeatureMode mode,
            string description,
            bool canRasterize = false) {
            switch (mode) {
                case UnsupportedFeatureMode.Error:
                    report.Add(
                        TranslationSeverity.Error,
                        feature,
                        $"The document contains {count} {description}, and the selected policy requires native preservation.",
                        code: codePrefix + ".UNSUPPORTED",
                        action: TranslationAction.Fail,
                        count: count);
                    break;
                case UnsupportedFeatureMode.WarnAndSkip:
                    report.Add(
                        TranslationSeverity.Warning,
                        feature,
                        $"The document contains {count} {description}; the current Google Docs exporter will skip them.",
                        code: codePrefix + ".SKIPPED",
                        action: TranslationAction.Skip,
                        count: count);
                    break;
                case UnsupportedFeatureMode.Flatten:
                    report.Add(
                        TranslationSeverity.Warning,
                        feature,
                        $"The document contains {count} {description}; a readable fallback notice will be inserted in the Google document.",
                        code: codePrefix + ".FLATTENED",
                        action: TranslationAction.Flatten,
                        count: count);
                    break;
                case UnsupportedFeatureMode.Rasterize when canRasterize:
                    report.Add(
                        TranslationSeverity.Warning,
                        feature,
                        $"The document contains {count} {description}; the OfficeIMO renderer will add a first-page PNG fallback through a temporary Drive lease.",
                        code: codePrefix + ".RASTERIZED",
                        action: TranslationAction.Rasterize,
                        count: count);
                    break;
                case UnsupportedFeatureMode.Rasterize:
                    report.Add(
                        TranslationSeverity.Error,
                        feature,
                        $"Rasterizing {description} requires InlineImageMode=TemporaryPublicDriveLease so Google Docs can fetch the rendered image safely.",
                        code: codePrefix + ".FALLBACK_UNAVAILABLE",
                        action: TranslationAction.Fail,
                        count: count);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(mode));
            }
        }
    }
}
