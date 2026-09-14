using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    internal static partial class PdfWordConverter {
        private static void ReportNonReconstructedLinks(PdfCore.PdfDocumentReadResult source, PdfToWordOptions options, ImportNavigationMap navigation) {
            int linkCount = source.Links.Count(link => !TryResolveWordLinkTarget(link, options, navigation, out _));
            if (linkCount == 0) return;

            AddWarning(
                options,
                "PdfLinkAnnotationNotReconstructed",
                "LinkAnnotation",
                "PDF link annotations that are remote, named viewer actions, unsafe, or unresolved are reported as diagnostics.",
                PdfCore.PdfConversionWarningSeverity.Information,
                OfficeConversionLossKind.Omission,
                new Dictionary<string, string> {
                    ["LinkCount"] = linkCount.ToString(CultureInfo.InvariantCulture)
                });
        }

        private static void ReportDocumentReconstructionBoundaries(PdfCore.PdfDocumentReadResult source, PdfToWordOptions options) {
            bool importsText = options.ImportHeadings || options.ImportParagraphs || options.ImportLists;
            bool reconstructsEditableContent = source.Pages.Any(page =>
                (importsText && page.TextBlocks.Count > 0) ||
                (options.ImportTables && page.Tables.Count > 0) ||
                ((options.ImportImages || options.IncludeImagePlaceholders) &&
                 page.Images.Any(image => PdfCore.PdfImagePlacementImportPolicy.HasVisiblePlacement(page, image)))) ||
                (options.IncludeFormFieldPlaceholders && source.FormFields.Count > 0);
            if (reconstructsEditableContent) {
                AddWarning(
                    options,
                    "PdfEditableLayoutReconstructed",
                    "Document",
                    "Editable Word output reconstructs document flow from PDF semantics; exact coordinates, pagination, grouping, and original authoring structure cannot be recovered reliably.",
                    PdfCore.PdfConversionWarningSeverity.Information,
                    OfficeConversionLossKind.Approximation);
            }
            if (!options.ImportHeadings) {
                ReportDisabledSemanticFamily(
                    options,
                    "PdfHeadingsNotImported",
                    "Document/Headings",
                    "PDF headings were not imported because ImportHeadings is false.",
                    "HeadingCount",
                    source.Pages.Sum(static page => page.Headings.Count));
            }
            if (!options.ImportParagraphs) {
                int paragraphCount = source.Pages.Sum(static page =>
                    page.Paragraphs.Count + page.TextBlocks.Count(static block =>
                        block.Kind is PdfCore.PdfLogicalElementKind.Header or
                            PdfCore.PdfLogicalElementKind.Footer or
                            PdfCore.PdfLogicalElementKind.Caption or
                            PdfCore.PdfLogicalElementKind.Footnote));
                ReportDisabledSemanticFamily(
                    options,
                    "PdfParagraphsNotImported",
                    "Document/Paragraphs",
                    "PDF paragraphs, headers, footers, captions, and footnotes were not imported because ImportParagraphs is false.",
                    "ParagraphLikeCount",
                    paragraphCount);
            }
            if (!options.ImportLists) {
                ReportDisabledSemanticFamily(
                    options,
                    "PdfListsNotImported",
                    "Document/Lists",
                    "PDF list items were not imported because ImportLists is false.",
                    "ListItemCount",
                    source.Pages.Sum(static page => page.ListItems.Count));
            }
            if (!importsText) {
                PdfCore.PdfTableExtractionScopeReport tableScope =
                    PdfCore.PdfLogicalTableAnalysis.AnalyzeExtractionScope(source);
                if (tableScope.NonTableTextBlockCount > 0) {
                    AddWarning(
                        options,
                        "PdfTextContentNotImported",
                        "Document/Text",
                        "Visible PDF text outside detected tables was not imported because every editable text option is disabled.",
                        PdfCore.PdfConversionWarningSeverity.Warning,
                        OfficeConversionLossKind.Omission,
                        new Dictionary<string, string> {
                            ["NonTableTextBlockCount"] = tableScope.NonTableTextBlockCount.ToString(CultureInfo.InvariantCulture)
                        });
                }
                if (tableScope.AnalysisTruncated) {
                    AddWarning(
                        options,
                        "PdfTableScopeAnalysisTruncated",
                        "Document/Text",
                        "Bounded table-scope analysis ended before every visible text block could be classified as represented by a table.",
                        PdfCore.PdfConversionWarningSeverity.Warning,
                        OfficeConversionLossKind.Approximation);
                }
            }
            if (!options.ImportTables) {
                ReportDisabledSemanticFamily(
                    options,
                    "PdfTablesNotImported",
                    "Document/Tables",
                    "Detected PDF tables were not imported because ImportTables is false.",
                    "TableCount",
                    source.Pages.Sum(static page => page.Tables.Count));
            }
            if (source.Outlines.Count > 0) {
                AddWarning(
                    options,
                    "PdfOutlineHierarchyNotReconstructed",
                    "Document/Outlines",
                    "PDF outline hierarchy is retained as diagnostic source metadata; semantic Word import reconstructs supported destinations and links, not the viewer outline tree.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> { ["OutlineCount"] = source.Outlines.Count.ToString(CultureInfo.InvariantCulture) });
            }
            if (source.TaggedContent != null) {
                AddWarning(
                    options,
                    "PdfTaggedStructureNotReconstructed",
                    "Document/StructTreeRoot",
                    "Readable PDF tagged-structure evidence informed the logical source model but was not copied as a Word accessibility structure tree.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> {
                        ["StructureElementCount"] = source.TaggedContent.StructureElementCount.ToString(CultureInfo.InvariantCulture),
                        ["MarkedContentReferenceCount"] = source.TaggedContent.MarkedContentReferenceCount.ToString(CultureInfo.InvariantCulture)
                    });
            }
            int optionalContentPageCount = source.Pages.Count(static page => page.HasOptionalContentUsage);
            if (optionalContentPageCount > 0) {
                AddWarning(
                    options,
                    "PdfOptionalContentGroupsFlattened",
                    "Document/OCProperties",
                    "PDF optional-content groups were flattened into the visible logical reconstruction; Word layer controls were not created.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> {
                        ["GroupCount"] = source.OptionalContentGroupCount.ToString(CultureInfo.InvariantCulture),
                        ["PageCount"] = optionalContentPageCount.ToString(CultureInfo.InvariantCulture)
                    });
            }
            if (source.CatalogActions.Count > 0 || source.OpenAction != null) {
                AddWarning(
                    options,
                    "PdfCatalogActionsNotReconstructed",
                    "Document/CatalogActions",
                    "PDF document open and catalog actions were not copied into the editable Word document.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> {
                        ["CatalogActionCount"] = source.CatalogActions.Count.ToString(CultureInfo.InvariantCulture),
                        ["HasOpenAction"] = (source.OpenAction != null).ToString(CultureInfo.InvariantCulture)
                    });
            }
        }

        private static void ReportDisabledSemanticFamily(
            PdfToWordOptions options,
            string code,
            string source,
            string message,
            string countName,
            int count) {
            if (count <= 0) return;
            AddWarning(
                options,
                code,
                source,
                message,
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission,
                new Dictionary<string, string> {
                    [countName] = count.ToString(CultureInfo.InvariantCulture)
                });
        }

        private static void ReportPageReconstructionBoundaries(PdfCore.PdfLogicalPage page, PdfToWordOptions options) {
            if (page.VectorPrimitiveCount > 0) {
                bool representedByImportedTableSemantics = options.ImportTables &&
                    page.Tables.Count > 0 &&
                    page.UnrepresentedVectorPrimitiveCount == 0;
                AddWarning(
                    options,
                    "PdfVectorGraphicsReconstructedSemantically",
                    "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Vectors",
                    "PDF vector primitives are not projected as editable Word shapes; only primitives proven to be represented table borders are treated as reconstructed semantics.",
                    representedByImportedTableSemantics
                        ? PdfCore.PdfConversionWarningSeverity.Information
                        : PdfCore.PdfConversionWarningSeverity.Warning,
                    representedByImportedTableSemantics
                        ? OfficeConversionLossKind.None
                        : OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> {
                        ["VectorPrimitiveCount"] = page.VectorPrimitiveCount.ToString(CultureInfo.InvariantCulture),
                        ["UnrepresentedVectorPrimitiveCount"] = page.UnrepresentedVectorPrimitiveCount.ToString(CultureInfo.InvariantCulture)
                    });
            }

            int annotationCount = page.Annotations.Count(static annotation =>
                !string.Equals(annotation.Subtype, "Link", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(annotation.Subtype, "Widget", StringComparison.OrdinalIgnoreCase));
            if (annotationCount > 0) {
                AddWarning(
                    options,
                    "PdfAnnotationsNotReconstructed",
                    "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Annotations",
                    "Non-link PDF annotations are retained in source diagnostics but are not reconstructed as editable Word comments or drawing objects.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> { ["AnnotationCount"] = annotationCount.ToString(CultureInfo.InvariantCulture) });
            }
            if (page.PageActions.Count > 0) {
                AddWarning(
                    options,
                    "PdfPageActionsNotReconstructed",
                    "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Actions",
                    "PDF page actions and chained actions are not copied into the editable Word document.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    new Dictionary<string, string> { ["PageActionCount"] = page.PageActions.Count.ToString(CultureInfo.InvariantCulture) });
            }
        }

        private static void CopyMetadata(PdfCore.PdfMetadata source, WordDocument target) {
            target.BuiltinDocumentProperties.Title = source.Title;
            target.BuiltinDocumentProperties.Creator = source.Author;
            target.BuiltinDocumentProperties.Subject = source.Subject;
            target.BuiltinDocumentProperties.Keywords = source.Keywords;
        }

        private static void AddWarning(
            PdfToWordOptions options,
            string code,
            string source,
            string message,
            PdfCore.PdfConversionWarningSeverity severity,
            IReadOnlyDictionary<string, string>? details = null) =>
            AddWarning(options, code, source, message, severity, lossKind: null, details);

        private static void AddWarning(
            PdfToWordOptions options,
            string code,
            string source,
            string message,
            PdfCore.PdfConversionWarningSeverity severity,
            OfficeConversionLossKind? lossKind,
            IReadOnlyDictionary<string, string>? details = null) {
            options.Report.Add(new PdfCore.PdfConversionWarning(
                ConverterName,
                code,
                source,
                message,
                severity,
                lossKind ?? (severity == PdfCore.PdfConversionWarningSeverity.Error
                    ? OfficeConversionLossKind.Failure
                    : severity == PdfCore.PdfConversionWarningSeverity.Information
                        ? OfficeConversionLossKind.None
                        : OfficeConversionLossKind.Approximation),
                details: details));
        }
    }
}
