using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenXml.Internal;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Finds metadata, personal information and hidden data, the way Office's Document Inspector does, without changing the workbook.</summary>
        public OfficeHiddenDataReport InspectHiddenData() => CollectHiddenData(null);

        /// <summary>Removes the selected categories of metadata and hidden data, then reports what was found and what happened to each item. Save the workbook to keep the result.</summary>
        /// <param name="options">Categories to remove; defaults to <see cref="OfficeHiddenDataKind.Default"/>, which keeps hidden sheets, rows, columns and names.</param>
        public OfficeHiddenDataReport RemoveHiddenData(OfficeHiddenDataOptions? options = null) {
            OfficeHiddenDataReport report = CollectHiddenData(options ?? new OfficeHiddenDataOptions());
            if (report.Changed) {
                DisqualifyDirectDataSetFastSaveState();
                MarkPackagePropertiesDirty();
                MarkApplicationPropertiesDirty();
                MarkPackageDirty();
            }
            return report;
        }

        private const string DataMashupNamespace = "http://schemas.microsoft.com/DataMashup";

        private OfficeHiddenDataReport CollectHiddenData(OfficeHiddenDataOptions? removal) {
            SpreadsheetDocument package = OpenXmlDocument;
            WorkbookPart workbookPart = package.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
            var collector = new OfficeHiddenDataCollector("Excel", removal);

            CollectExcelComments(collector);
            CollectExcelHiddenContent(collector);
            CollectExcelIdentities(workbookPart, collector);
            OfficeOpenXmlHiddenData.InspectDocumentProperties(package, collector);
            OfficeOpenXmlHiddenData.InspectCustomProperties(package, collector, name => RemoveCustomDocumentProperty(name));
            CollectExcelExternalLinks(workbookPart, collector);
            CollectExcelPivotCacheSources(workbookPart, collector);
            OfficeOpenXmlHiddenData.InspectCustomXml(workbookPart, workbookPart.CustomXmlParts, collector, part =>
                OfficeOpenXmlHiddenData.ReadRootNamespace(part) == DataMashupNamespace
                    ? "Power Query stores the workbook's queries here; delete the queries in Excel to remove it."
                    : null);
            OfficeOpenXmlHiddenData.InspectThumbnail(package, collector);
            OfficeOpenXmlHiddenData.InspectImagesAndProvenance(package, collector);
            OfficeOpenXmlHiddenData.InspectPrinterSettings(package, collector);
            return collector.Build();
        }

        private void CollectExcelComments(OfficeHiddenDataCollector collector) {
            foreach (ExcelSheet sheet in Sheets) {
                foreach (ExcelThreadedCommentInfo comment in sheet.GetThreadedComments()) {
                    if (comment.ParentId != null || comment.Id == null) continue;
                    string id = comment.Id;
                    collector.Add(OfficeHiddenDataKind.Comments, sheet.Name + "!" + comment.CellReference,
                        "Threaded comment by " + (comment.Author ?? "unknown author"), comment.Text,
                        () => sheet.RemoveThreadedComment(id, removeReplies: true));
                }

                foreach (ExcelCommentInfo note in sheet.GetComments()) {
                    string reference = note.CellReference;
                    collector.Add(OfficeHiddenDataKind.Comments, sheet.Name + "!" + reference,
                        "Note by " + (note.Author ?? "unknown author"), note.Text, () => sheet.ClearComment(reference));
                }
            }
        }

        private void CollectExcelHiddenContent(OfficeHiddenDataCollector collector) {
            ExcelWorkbookInfo snapshot = Inspect();
            var hiddenSheetNames = new HashSet<string>(snapshot.Worksheets.Where(sheet => sheet.Hidden).Select(sheet => sheet.Name), StringComparer.OrdinalIgnoreCase);

            foreach (ExcelNamedRangeInfo name in ListNamedRanges(includeBuiltIn: false, includeHidden: true)) {
                bool pointsAtHiddenSheet = hiddenSheetNames.Any(sheetName => ReferencesSheet(name.Reference, sheetName));
                if (!name.Hidden && !pointsAtHiddenSheet) continue;
                string description = name.Hidden ? "Hidden defined name '" + name.Name + "'" : "Defined name '" + name.Name + "' pointing to a hidden sheet";
                ExcelSheet? scope = name.SheetName == null ? null : Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, name.SheetName, StringComparison.OrdinalIgnoreCase));
                string definedName = name.Name;
                collector.Add(OfficeHiddenDataKind.HiddenContent, name.SheetName ?? "Workbook", description, name.Reference,
                    () => RemoveNamedRange(definedName, scope, save: false));
            }

            foreach (ExcelWorksheetInfo worksheet in snapshot.Worksheets.Where(sheet => !sheet.Hidden)) {
                ExcelSheet sheet = this[worksheet.Name];
                foreach (int row in worksheet.Rows.Where(row => row.Hidden).Select(row => row.Index).OrderByDescending(index => index).ToList()) {
                    collector.Add(OfficeHiddenDataKind.HiddenContent, worksheet.Name + "!" + row.ToString(CultureInfo.InvariantCulture) + ":" + row.ToString(CultureInfo.InvariantCulture),
                        "Hidden row", null, () => sheet.DeleteRows(row));
                }
                foreach (ExcelColumnInfo column in worksheet.Columns.Where(column => column.Hidden).OrderByDescending(column => column.StartIndex).ToList()) {
                    int first = column.StartIndex;
                    int count = column.EndIndex - column.StartIndex + 1;
                    if (count > 1024) {
                        collector.Add(OfficeHiddenDataKind.HiddenContent, worksheet.Name, "Hidden column range " + ColumnRange(first, column.EndIndex), null,
                            null, "Hidden column ranges this wide usually hide the unused part of the sheet; they are left in place.");
                        continue;
                    }
                    collector.Add(OfficeHiddenDataKind.HiddenContent, worksheet.Name + "!" + ColumnRange(first, column.EndIndex),
                        count == 1 ? "Hidden column" : "Hidden columns", null, () => sheet.DeleteColumns(first, count));
                }
            }

            foreach (ExcelWorksheetInfo worksheet in snapshot.Worksheets.Where(sheet => sheet.Hidden)) {
                string name = worksheet.Name;
                collector.Add(OfficeHiddenDataKind.HiddenContent, name,
                    worksheet.VisibilityState == "veryHidden" ? "Very hidden worksheet" : "Hidden worksheet", null, () => RemoveWorksheet(name));
            }
        }

        // Matches the sheet as a whole reference token, so a hidden "Data" doesn't match "RawData!" or "[1]Data!".
        private static bool ReferencesSheet(string reference, string sheetName) =>
            Regex.IsMatch(reference,
                @"(?<![\w.'\]])(?:" + Regex.Escape(sheetName) + "|'" + Regex.Escape(sheetName.Replace("'", "''")) + "')!",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static string ColumnRange(int first, int last) =>
            ExcelA1.ColumnIndexToLetters(first) + ":" + ExcelA1.ColumnIndexToLetters(last);

        private static void CollectExcelIdentities(WorkbookPart workbookPart, OfficeHiddenDataCollector collector) {
            bool threadedCommentsInUse = workbookPart.WorksheetParts.Any(worksheet =>
                worksheet.WorksheetThreadedCommentsParts.Any(comments => comments.ThreadedComments?.HasChildren == true));
            foreach (WorkbookPersonPart part in workbookPart.WorkbookPersonParts.ToList()) {
                if (!threadedCommentsInUse) {
                    string names = string.Join(", ", part.PersonList?.Elements<Threaded.Person>().Select(person => person.DisplayName?.Value) ?? Array.Empty<string?>());
                    collector.Add(OfficeHiddenDataKind.PersonalIdentities, part.Uri.ToString(), "Threaded comment authors no comment refers to", names,
                        () => workbookPart.DeletePart(part));
                    continue;
                }
                foreach (Threaded.Person person in part.PersonList?.Elements<Threaded.Person>().ToList() ?? new List<Threaded.Person>()) {
                    string? name = person.DisplayName?.Value;
                    if (collector.Removal != null && name == collector.Removal.AnonymousAuthor && person.UserId == null) continue;
                    collector.Add(OfficeHiddenDataKind.PersonalIdentities, part.Uri.ToString(), "Threaded comment author", name, () => {
                        person.DisplayName = collector.Removal!.AnonymousAuthor;
                        person.UserId = null;
                        person.ProviderId = null;
                    }, removedOutcome: OfficeHiddenDataOutcome.Anonymized);
                }
            }

            foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts) {
                Authors? authors = worksheetPart.WorksheetCommentsPart?.Comments?.Authors;
                foreach (Author author in authors?.Elements<Author>().ToList() ?? new List<Author>()) {
                    if (string.IsNullOrEmpty(author.Text)) continue;
                    // Excel links the legacy note behind a threaded comment through a "tc={id}" author; it isn't a name.
                    if (author.Text!.StartsWith("tc=", StringComparison.OrdinalIgnoreCase)) continue;
                    if (collector.Removal != null && author.Text == collector.Removal.AnonymousAuthor) continue;
                    collector.Add(OfficeHiddenDataKind.PersonalIdentities, worksheetPart.WorksheetCommentsPart!.Uri.ToString(), "Note author", author.Text,
                        () => author.Text = collector.Removal!.AnonymousAuthor, removedOutcome: OfficeHiddenDataOutcome.Anonymized);
                }
            }
        }

        private static void CollectExcelExternalLinks(WorkbookPart workbookPart, OfficeHiddenDataCollector collector) {
            foreach (ExternalWorkbookPart part in workbookPart.ExternalWorkbookParts) {
                string location = part.Uri.ToString();
                ExternalBook? book = part.ExternalLink?.GetFirstChild<ExternalBook>();
                string? path = part.ExternalRelationships.FirstOrDefault(relationship => relationship.Id == book?.Id?.Value)?.Uri.OriginalString;
                collector.Add(OfficeHiddenDataKind.ExternalReferences, location, "External workbook link", path, null,
                    "Formulas refer to this link; break the links in Excel to remove the path.");
                SheetDataSet? cache = book?.SheetDataSet;
                if (cache != null && cache.HasChildren) {
                    collector.Add(OfficeHiddenDataKind.ExternalReferences, location, "Cached values from the external workbook", path, cache.Remove);
                }
            }
        }

        private static void CollectExcelPivotCacheSources(WorkbookPart workbookPart, OfficeHiddenDataCollector collector) {
            foreach (PivotTableCacheDefinitionPart part in workbookPart.PivotTableCacheDefinitionParts) {
                CacheSource? source = part.PivotCacheDefinition?.CacheSource;
                WorksheetSource? worksheetSource = source?.WorksheetSource;
                string? relationshipId = worksheetSource?.Id?.Value;
                if (relationshipId == null) continue;
                string? path = part.ExternalRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId)?.Uri.OriginalString;
                if (path == null) continue;
                collector.Add(OfficeHiddenDataKind.ExternalReferences, part.Uri.ToString(), "Pivot cache source workbook", path, null,
                    "The pivot table refreshes from this source; change its data source in Excel to remove the path.");
            }
        }
    }
}
