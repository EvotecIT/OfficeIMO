using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Removes empty containers and orphaned references on this worksheet to prevent Excel repairs.
        /// </summary>
        internal void Preflight() {
            WriteLock(() => {
                var ws = WorksheetRoot;

                CleanupHyperlinkArtifacts();

                // Remove empty Hyperlinks
                var links = ws.Elements<Hyperlinks>().FirstOrDefault();
                if (links != null && !links.Elements<Hyperlink>().Any()) {
                    ws.RemoveChild(links);
                }

                // Remove empty MergeCells
                var merges = ws.Elements<MergeCells>().FirstOrDefault();
                if (merges != null && !merges.Elements<MergeCell>().Any()) {
                    ws.RemoveChild(merges);
                }

                // Remove empty DataValidations containers
                var dataValidations = ws.Elements<DataValidations>().FirstOrDefault();
                if (dataValidations != null) {
                    var validationCount = dataValidations.Elements<DataValidation>().Count();
                    if (validationCount == 0) {
                        ws.RemoveChild(dataValidations);
                    } else {
                        dataValidations.SetAttribute(new OpenXmlAttribute("count", string.Empty, validationCount.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                    }
                }

                // Remove empty IgnoredErrors containers
                var ignoredErrors = ws.Elements<IgnoredErrors>().FirstOrDefault();
                if (ignoredErrors != null) {
                    var errorCount = ignoredErrors.Elements<IgnoredError>().Count();
                    if (errorCount == 0) {
                        ws.RemoveChild(ignoredErrors);
                    } else {
                        ignoredErrors.SetAttribute(new OpenXmlAttribute("count", string.Empty, errorCount.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                    }
                }

                // Remove empty CustomSheetViews containers
                var customSheetViews = ws.Elements<CustomSheetViews>().FirstOrDefault();
                if (customSheetViews != null && !customSheetViews.Elements<CustomSheetView>().Any()) {
                    ws.RemoveChild(customSheetViews);
                }

                CleanupSheetViewArtifacts();
                CleanupPrintArtifacts();

                // Remove empty or malformed ConditionalFormatting entries
                foreach (var conditional in ws.Elements<ConditionalFormatting>().ToList()) {
                    foreach (var rule in conditional.Elements<ConditionalFormattingRule>().ToList()) {
                        if (rule.Type == null) {
                            rule.Remove();
                        }
                    }

                    bool hasRules = conditional.Elements<ConditionalFormattingRule>().Any();
                    bool hasRanges = !string.IsNullOrWhiteSpace(conditional.SequenceOfReferences?.InnerText);
                    if (!hasRules || !hasRanges) {
                        conditional.Remove();
                    }
                }

                int nextConditionalPriority = 1;
                foreach (var rule in ws.Elements<ConditionalFormatting>().SelectMany(cf => cf.Elements<ConditionalFormattingRule>())) {
                    if (rule.Priority?.Value != nextConditionalPriority) {
                        rule.Priority = nextConditionalPriority;
                    }
                    nextConditionalPriority++;
                }

                CleanupProtectionArtifacts();
                CleanupTableArtifacts();
                CleanupAutoFilterArtifacts();
                CleanupCommentArtifacts();
                CleanupHeaderFooterPictureArtifacts();
                CleanupWorksheetDrawingArtifacts();

                // Drop orphaned Drawing reference
                var drawing = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
                if (drawing?.Id?.Value is string dId) {
                    try { _worksheetPart.GetPartById(dId); } catch (ArgumentOutOfRangeException) { ws.RemoveChild(drawing); }
                }

                // Drop orphaned LegacyDrawingHeaderFooter reference
                var legacy = ws.GetFirstChild<LegacyDrawingHeaderFooter>();
                if (legacy?.Id?.Value is string lId) {
                    try { _worksheetPart.GetPartById(lId); } catch (ArgumentOutOfRangeException) { ws.RemoveChild(legacy); }
                }

                ws.Save();
            });
        }
    }
}

