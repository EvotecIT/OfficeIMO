using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options that control safe workbook repair orchestration.
    /// </summary>
    public sealed class ExcelWorkbookRepairOptions {
        /// <summary>Repair duplicate, invalid, and broken workbook defined names.</summary>
        public bool DefinedNames { get; set; } = true;

        /// <summary>Normalize worksheet table metadata, table relationships, and table column definitions.</summary>
        public bool Tables { get; set; } = true;

        /// <summary>Normalize worksheet view metadata such as frozen panes and selections.</summary>
        public bool SheetViews { get; set; } = true;

        /// <summary>Normalize print metadata such as page breaks, margins, and page scale.</summary>
        public bool PrintSettings { get; set; } = true;

        /// <summary>Normalize worksheet protection metadata and protected range declarations.</summary>
        public bool Protection { get; set; } = true;

        /// <summary>Normalize AutoFilter metadata when the worksheet exposes a stale or invalid filter range.</summary>
        public bool AutoFilters { get; set; } = true;

        /// <summary>Normalize drawing relationships and worksheet drawing anchors.</summary>
        public bool Drawings { get; set; } = true;

        /// <summary>Normalize hyperlinks and comment artifacts.</summary>
        public bool LinksAndComments { get; set; } = true;

        /// <summary>Remove stale calculation chains and force recalculation on open when formulas exist.</summary>
        public bool Calculation { get; set; } = true;

        /// <summary>Normalize workbook views, style, and shared-string artifacts.</summary>
        public bool WorkbookArtifacts { get; set; } = true;

        /// <summary>Save the workbook after repairs are applied.</summary>
        public bool Save { get; set; } = true;
    }

    /// <summary>
    /// One safe repair action applied to a workbook.
    /// </summary>
    public sealed class ExcelWorkbookRepairAction {
        internal ExcelWorkbookRepairAction(string category, string message, string? sheetName = null) {
            Category = category;
            Message = message;
            SheetName = sheetName;
        }

        /// <summary>Repair category such as DefinedName, Table, View, or Drawing.</summary>
        public string Category { get; }

        /// <summary>Human-readable repair action summary.</summary>
        public string Message { get; }

        /// <summary>Worksheet name when the repair was worksheet scoped.</summary>
        public string? SheetName { get; }
    }

    /// <summary>
    /// Report returned after safe workbook repair orchestration.
    /// </summary>
    public sealed class ExcelWorkbookRepairReport {
        internal ExcelWorkbookRepairReport(IReadOnlyList<ExcelWorkbookRepairAction> actions, ExcelWorkbookDiagnosticReport before, ExcelWorkbookDiagnosticReport after) {
            Actions = actions;
            Before = before;
            After = after;
        }

        /// <summary>Actions attempted by the repair pass.</summary>
        public IReadOnlyList<ExcelWorkbookRepairAction> Actions { get; }

        /// <summary>Diagnostic report captured before repairs.</summary>
        public ExcelWorkbookDiagnosticReport Before { get; }

        /// <summary>Diagnostic report captured after repairs.</summary>
        public ExcelWorkbookDiagnosticReport After { get; }

        /// <summary>Number of repair actions attempted.</summary>
        public int ActionCount => Actions.Count;
    }

    /// <summary>
    /// Options for workbook comparison depth.
    /// </summary>
    public sealed class ExcelWorkbookDiffOptions {
        /// <summary>Maximum differences to report.</summary>
        public int MaxDifferences { get; set; } = 200;

        /// <summary>Compare visible cell values and formulas.</summary>
        public bool CompareCells { get; set; } = true;

        /// <summary>Compare cell style indexes in used cells.</summary>
        public bool CompareCellStyles { get; set; } = true;

        /// <summary>Compare workbook and sheet-scoped defined names.</summary>
        public bool CompareNamedRanges { get; set; } = true;

        /// <summary>Compare table names, ranges, and header-row state.</summary>
        public bool CompareTables { get; set; } = true;

        /// <summary>Compare worksheet validations and AutoFilter/view state.</summary>
        public bool CompareWorksheetMetadata { get; set; } = true;

        /// <summary>Compare legacy and threaded comments.</summary>
        public bool CompareComments { get; set; } = true;
    }

    /// <summary>
    /// Comment and threaded-comment awareness report.
    /// </summary>
    public sealed class ExcelWorkbookCommentReport {
        internal ExcelWorkbookCommentReport(IReadOnlyList<ExcelCommentRecord> comments, IReadOnlyList<ExcelThreadedCommentSnapshot> threadedComments, IReadOnlyList<ExcelWorkbookDiagnosticIssue> issues) {
            Comments = comments;
            ThreadedComments = threadedComments;
            Issues = issues;
        }

        /// <summary>Legacy worksheet notes across the workbook.</summary>
        public IReadOnlyList<ExcelCommentRecord> Comments { get; }

        /// <summary>Threaded comments preserved in the workbook package.</summary>
        public IReadOnlyList<ExcelThreadedCommentSnapshot> ThreadedComments { get; }

        /// <summary>Comment metadata issues, such as missing authors or empty comment text.</summary>
        public IReadOnlyList<ExcelWorkbookDiagnosticIssue> Issues { get; }

        /// <summary>Legacy comment count.</summary>
        public int CommentCount => Comments.Count;

        /// <summary>Threaded comment count.</summary>
        public int ThreadedCommentCount => ThreadedComments.Count;

        /// <summary>True when any comment or threaded comment exists.</summary>
        public bool HasComments => CommentCount > 0 || ThreadedCommentCount > 0;
    }

    /// <summary>
    /// Legacy worksheet note projected with workbook-level sheet context.
    /// </summary>
    public sealed class ExcelCommentRecord {
        internal ExcelCommentRecord(string sheetName, string cellReference, string? author, string text) {
            SheetName = sheetName;
            CellReference = cellReference;
            Author = author;
            Text = text;
        }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; }

        /// <summary>A1 cell reference.</summary>
        public string CellReference { get; }

        /// <summary>Comment author display name, when available.</summary>
        public string? Author { get; }

        /// <summary>Comment text.</summary>
        public string Text { get; }
    }

    /// <summary>
    /// Template binding diagnostics for a workbook or sheet template.
    /// </summary>
    public sealed class ExcelTemplateBindingReport {
        internal ExcelTemplateBindingReport(ExcelTemplateInspection inspection) {
            Inspection = inspection;
        }

        /// <summary>Underlying marker inspection.</summary>
        public ExcelTemplateInspection Inspection { get; }

        /// <summary>Total marker occurrences.</summary>
        public int TotalMarkers => Inspection.TotalMarkers;

        /// <summary>Distinct missing marker names.</summary>
        public IReadOnlyList<string> MissingMarkerNames => Inspection.MissingMarkerNames;

        /// <summary>True when the supplied bindings satisfy all markers.</summary>
        public bool Passed => Inspection.HasBindingInfo && Inspection.MissingMarkerNames.Count == 0;

        /// <summary>Markdown table describing all discovered markers and binding state.</summary>
        public string Markdown => Inspection.ToMarkdown();
    }

    /// <summary>
    /// Safe metadata options for workbook connection and query-table package parts.
    /// </summary>
    public sealed class ExcelPowerQueryMetadataOptions {
        /// <summary>Connection name shown by Excel-compatible applications.</summary>
        public string Name { get; set; } = "OfficeIMOQuery";

        /// <summary>Optional worksheet name that should own query-table metadata.</summary>
        public string? WorksheetName { get; set; }

        /// <summary>Optional query-table name. Defaults to the connection name with "Table" suffix.</summary>
        public string? QueryTableName { get; set; }

        /// <summary>Connection description stored in package metadata.</summary>
        public string? Description { get; set; }

        /// <summary>Power Query M expression stored as metadata for Excel to own and execute.</summary>
        public string? CommandText { get; set; }

        /// <summary>Request refresh-on-open metadata on the authored connection.</summary>
        public bool RefreshOnOpen { get; set; }
    }

    /// <summary>
    /// Result from authoring safe workbook query metadata.
    /// </summary>
    public sealed class ExcelPowerQueryMetadataResult {
        internal ExcelPowerQueryMetadataResult(string connectionName, uint connectionId, string? queryTableName, bool addedWorkbookConnection, bool addedWorksheetQueryTable, bool refreshOnOpen) {
            ConnectionName = connectionName;
            ConnectionId = connectionId;
            QueryTableName = queryTableName;
            AddedWorkbookConnection = addedWorkbookConnection;
            AddedWorksheetQueryTable = addedWorksheetQueryTable;
            RefreshOnOpen = refreshOnOpen;
        }

        /// <summary>Connection name stored in metadata.</summary>
        public string ConnectionName { get; }

        /// <summary>Workbook connection id used by authored query-table metadata.</summary>
        public uint ConnectionId { get; }

        /// <summary>Query-table name stored in worksheet metadata, when requested.</summary>
        public string? QueryTableName { get; }

        /// <summary>True when workbook connection metadata was authored.</summary>
        public bool AddedWorkbookConnection { get; }

        /// <summary>True when worksheet query-table metadata was authored.</summary>
        public bool AddedWorksheetQueryTable { get; }

        /// <summary>True when refresh-on-open metadata was requested.</summary>
        public bool RefreshOnOpen { get; }
    }

    public partial class ExcelDocument {
        /// <summary>
        /// Runs safe workbook repairs using OfficeIMO's existing cleanup engines, then returns before/after diagnostics.
        /// </summary>
        public ExcelWorkbookRepairReport RepairWorkbook(ExcelWorkbookRepairOptions? options = null) {
            options ??= new ExcelWorkbookRepairOptions();
            ExcelWorkbookDiagnosticReport before = RunWorkbookDoctor(new ExcelWorkbookDoctorOptions { ValidateOpenXml = false });
            var actions = new List<ExcelWorkbookRepairAction>();

            if (options.DefinedNames) {
                RepairDefinedNames(save: false);
                CleanupDefinedNameArtifacts(includeAggressiveRepairs: true, save: false);
                actions.Add(new ExcelWorkbookRepairAction("DefinedName", "Normalized duplicate, hidden, and broken defined-name artifacts."));
            }

            foreach (ExcelSheet sheet in Sheets) {
                if (options.Tables) {
                    sheet.CleanupTableArtifacts();
                    actions.Add(new ExcelWorkbookRepairAction("Table", "Normalized table relationships, identifiers, names, and columns.", sheet.Name));
                }
                if (options.SheetViews) {
                    sheet.CleanupSheetViewArtifacts();
                    actions.Add(new ExcelWorkbookRepairAction("View", "Normalized worksheet view, pane, and selection metadata.", sheet.Name));
                }
                if (options.PrintSettings) {
                    sheet.CleanupPrintArtifacts();
                    actions.Add(new ExcelWorkbookRepairAction("Print", "Normalized print settings, page breaks, margins, and scale.", sheet.Name));
                }
                if (options.Protection) {
                    sheet.CleanupProtectionArtifacts();
                    actions.Add(new ExcelWorkbookRepairAction("Protection", "Normalized worksheet protection and protected ranges.", sheet.Name));
                }
                if (options.AutoFilters) {
                    sheet.CleanupAutoFilterArtifacts();
                    actions.Add(new ExcelWorkbookRepairAction("AutoFilter", "Normalized worksheet AutoFilter artifacts.", sheet.Name));
                }
                if (options.Drawings) {
                    sheet.CleanupWorksheetDrawingArtifacts();
                    sheet.CleanupHeaderFooterPictureArtifacts();
                    actions.Add(new ExcelWorkbookRepairAction("Drawing", "Normalized drawing, image, and header/footer picture artifacts.", sheet.Name));
                }
                if (options.LinksAndComments) {
                    sheet.CleanupHyperlinkArtifacts();
                    actions.Add(new ExcelWorkbookRepairAction("Link", "Normalized hyperlink artifacts.", sheet.Name));
                }
            }

            if (options.Calculation) {
                CleanupCalculationArtifacts(save: false, ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
                actions.Add(new ExcelWorkbookRepairAction("Calculation", "Removed stale calculation chains and requested recalculation on open."));
            }
            if (options.WorkbookArtifacts) {
                CleanupWorkbookViewArtifacts(save: false);
                CleanupStyleAndSharedStringArtifacts(save: false);
                actions.Add(new ExcelWorkbookRepairAction("Workbook", "Normalized workbook views, styles, and shared strings."));
            }

            if (actions.Count > 0) {
                MarkPackageDirty();
                if (options.Save && CanSaveToDefaultDestination()) {
                    Save();
                }
            }

            ExcelWorkbookDiagnosticReport after = RunWorkbookDoctor(new ExcelWorkbookDoctorOptions { ValidateOpenXml = false });
            return new ExcelWorkbookRepairReport(actions, before, after);
        }

        /// <summary>
        /// Compares this workbook with another workbook using configurable structural, value, style, table, comment, and metadata checks.
        /// </summary>
        public ExcelWorkbookDiffReport CompareWorkbook(ExcelDocument other, ExcelWorkbookDiffOptions options) {
            if (other == null) throw new ArgumentNullException(nameof(other));
            if (options == null) throw new ArgumentNullException(nameof(options));
            int maxDifferences = Math.Max(1, options.MaxDifferences);
            var differences = new List<ExcelWorkbookDifference>();
            var otherSheets = other.Sheets.ToDictionary(sheet => sheet.Name, StringComparer.OrdinalIgnoreCase);
            bool compareSnapshots = options.CompareTables || options.CompareWorksheetMetadata || options.CompareComments;
            ExcelWorkbookSnapshot? leftSnapshot = compareSnapshots ? CreateInspectionSnapshot() : null;
            ExcelWorkbookSnapshot? rightSnapshot = compareSnapshots ? other.CreateInspectionSnapshot() : null;

            foreach (ExcelSheet sheet in Sheets) {
                if (!otherSheets.TryGetValue(sheet.Name, out ExcelSheet? otherSheet)) {
                    AddDifference(differences, maxDifferences, new ExcelWorkbookDifference("Sheet", "Sheet exists only in left workbook.", sheet.Name));
                    continue;
                }

                if (options.CompareCells) {
                    CompareSheetValues(sheet, otherSheet, differences, maxDifferences);
                }
                if (options.CompareCellStyles) {
                    CompareCellStyles(sheet, otherSheet, differences, maxDifferences);
                }
                if (compareSnapshots) {
                    CompareSheetSnapshots(leftSnapshot!, rightSnapshot!, sheet.Name, differences, maxDifferences, options);
                }
                if (differences.Count >= maxDifferences) break;
            }

            var leftNames = new HashSet<string>(Sheets.Select(sheet => sheet.Name), StringComparer.OrdinalIgnoreCase);
            foreach (ExcelSheet sheet in other.Sheets.Where(sheet => !leftNames.Contains(sheet.Name))) {
                AddDifference(differences, maxDifferences, new ExcelWorkbookDifference("Sheet", "Sheet exists only in right workbook.", sheet.Name));
            }

            if (options.CompareNamedRanges) {
                CompareNamedRanges(other, differences, maxDifferences);
            }

            return new ExcelWorkbookDiffReport(differences);
        }

        private bool CanSaveToDefaultDestination()
            => !string.IsNullOrEmpty(FilePath) || _sourceStream != null;

        /// <summary>
        /// Audits legacy notes and threaded comments across the workbook.
        /// </summary>
        public ExcelWorkbookCommentReport InspectComments() {
            var comments = new List<ExcelCommentRecord>();
            var issues = new List<ExcelWorkbookDiagnosticIssue>();
            foreach (ExcelSheet sheet in Sheets) {
                foreach (ExcelCommentInfo comment in sheet.GetComments()) {
                    comments.Add(new ExcelCommentRecord(sheet.Name, comment.CellReference, comment.Author, comment.Text));
                    if (string.IsNullOrWhiteSpace(comment.Author)) {
                        issues.Add(new ExcelWorkbookDiagnosticIssue("Comment", ExcelFindingSeverity.Warning, "Legacy comment has no author.", sheet.Name, comment.CellReference));
                    }
                    if (string.IsNullOrWhiteSpace(comment.Text)) {
                        issues.Add(new ExcelWorkbookDiagnosticIssue("Comment", ExcelFindingSeverity.Warning, "Legacy comment has no text.", sheet.Name, comment.CellReference));
                    }
                }
            }

            var threaded = CreateInspectionSnapshot().Worksheets.SelectMany(sheet => sheet.ThreadedComments).ToList();
            foreach (ExcelThreadedCommentSnapshot comment in threaded) {
                if (string.IsNullOrWhiteSpace(comment.Author)) {
                    issues.Add(new ExcelWorkbookDiagnosticIssue("ThreadedComment", ExcelFindingSeverity.Warning, "Threaded comment has no resolved author.", comment.SheetName, comment.CellReference));
                }
                if (!string.IsNullOrWhiteSpace(comment.ParentId) && threaded.All(parent => !string.Equals(parent.Id, comment.ParentId, StringComparison.OrdinalIgnoreCase))) {
                    issues.Add(new ExcelWorkbookDiagnosticIssue("ThreadedComment", ExcelFindingSeverity.Warning, "Threaded comment reply references a missing parent.", comment.SheetName, comment.CellReference));
                }
            }

            return new ExcelWorkbookCommentReport(comments, threaded, issues);
        }

        /// <summary>
        /// Validates all workbook template markers against a dictionary of values before applying the template.
        /// </summary>
        public ExcelTemplateBindingReport ValidateTemplateBindings(IDictionary<string, object?> values) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            return new ExcelTemplateBindingReport(InspectTemplate(values));
        }

        /// <summary>
        /// Validates all workbook template markers against an object model before applying the template.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("Object-model template inspection walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        public ExcelTemplateBindingReport ValidateTemplateBindings(object model) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return new ExcelTemplateBindingReport(InspectTemplate(model));
        }

        /// <summary>
        /// Authors safe connection/query-table metadata for Excel-compatible applications to own and refresh.
        /// OfficeIMO writes metadata only; it does not execute Power Query M or contact external systems.
        /// </summary>
        public ExcelPowerQueryMetadataResult AddPowerQueryMetadata(ExcelPowerQueryMetadataOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            string name = string.IsNullOrWhiteSpace(options.Name) ? "OfficeIMOQuery" : options.Name.Trim();
            string commandText = options.CommandText ?? string.Empty;
            string? worksheetName = null;
            if (!string.IsNullOrWhiteSpace(options.WorksheetName)) {
                worksheetName = this[options.WorksheetName!].Name;
            }

            uint connectionId = GetNextPowerQueryConnectionId();
            AddWorkbookConnectionMetadata(BuildConnectionMetadataXml(name, options.Description, commandText, options.RefreshOnOpen, connectionId));
            bool addedQueryTable = false;
            string? queryTableName = null;
            if (worksheetName != null) {
                queryTableName = NormalizeQueryTableName(options.QueryTableName, name);
                AddWorksheetQueryTableMetadata(worksheetName, BuildQueryTableMetadataXml(queryTableName, connectionId));
                addedQueryTable = true;
            }

            return new ExcelPowerQueryMetadataResult(name, connectionId, queryTableName, addedWorkbookConnection: true, addedQueryTable, options.RefreshOnOpen);
        }

        private static void CompareCellStyles(ExcelSheet left, ExcelSheet right, ICollection<ExcelWorkbookDifference> differences, int maxDifferences) {
            var rightCells = (right.WorksheetPart.Worksheet?.Descendants<Cell>() ?? Enumerable.Empty<Cell>())
                .Where(cell => !string.IsNullOrWhiteSpace(cell.CellReference?.Value))
                .ToDictionary(cell => cell.CellReference!.Value!, StringComparer.OrdinalIgnoreCase);
            foreach (Cell leftCell in left.WorksheetPart.Worksheet?.Descendants<Cell>() ?? Enumerable.Empty<Cell>()) {
                if (differences.Count >= maxDifferences) break;
                string? address = leftCell.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(address)) continue;
                rightCells.TryGetValue(address!, out Cell? rightCell);
                rightCells.Remove(address!);
                string leftStyle = leftCell.StyleIndex?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                string rightStyle = rightCell?.StyleIndex?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                if (!string.Equals(leftStyle, rightStyle, StringComparison.Ordinal)) {
                    AddDifference(differences, maxDifferences, new ExcelWorkbookDifference("CellStyle", "Cell style index differs.", left.Name, address, leftStyle, rightStyle));
                }
            }

            foreach (KeyValuePair<string, Cell> rightOnly in rightCells) {
                if (differences.Count >= maxDifferences) break;
                string rightStyle = rightOnly.Value.StyleIndex?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                if (!string.IsNullOrEmpty(rightStyle)) {
                    AddDifference(differences, maxDifferences, new ExcelWorkbookDifference("CellStyle", "Cell style index differs.", left.Name, rightOnly.Key, string.Empty, rightStyle));
                }
            }
        }

        private static void CompareSheetSnapshots(ExcelWorkbookSnapshot leftSnapshot, ExcelWorkbookSnapshot rightSnapshot, string sheetName, ICollection<ExcelWorkbookDifference> differences, int maxDifferences, ExcelWorkbookDiffOptions options) {
            ExcelWorksheetSnapshot? left = leftSnapshot.Worksheets.FirstOrDefault(sheet => string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            ExcelWorksheetSnapshot? right = rightSnapshot.Worksheets.FirstOrDefault(sheet => string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (left == null || right == null) return;

            if (options.CompareTables) {
                CompareKeyed(left.Tables.Select(table => $"{table.Name}|{table.A1Range}|{table.HasHeaderRow}"), right.Tables.Select(table => $"{table.Name}|{table.A1Range}|{table.HasHeaderRow}"), "Table", sheetName, differences, maxDifferences);
            }
            if (options.CompareWorksheetMetadata) {
                CompareScalar("WorksheetView", "Used range differs.", sheetName, left.UsedRangeA1, right.UsedRangeA1, differences, maxDifferences);
                CompareScalar("WorksheetView", "Gridline state differs.", sheetName, left.ShowGridlines.ToString(), right.ShowGridlines.ToString(), differences, maxDifferences);
                CompareScalar("WorksheetView", "Frozen pane state differs.", sheetName, $"{left.FrozenRowCount},{left.FrozenColumnCount}", $"{right.FrozenRowCount},{right.FrozenColumnCount}", differences, maxDifferences);
                CompareKeyed(left.Validations.Select(validation => $"{validation.Type}|{validation.Operator}|{string.Join(",", validation.A1Ranges)}"), right.Validations.Select(validation => $"{validation.Type}|{validation.Operator}|{string.Join(",", validation.A1Ranges)}"), "DataValidation", sheetName, differences, maxDifferences);
            }
            if (options.CompareComments) {
                CompareKeyed(left.Cells.Where(cell => cell.Comment != null).Select(cell => $"{A1.CellReference(cell.Row, cell.Column)}|{cell.Comment!.Author}|{cell.Comment.Text}"), right.Cells.Where(cell => cell.Comment != null).Select(cell => $"{A1.CellReference(cell.Row, cell.Column)}|{cell.Comment!.Author}|{cell.Comment.Text}"), "Comment", sheetName, differences, maxDifferences);
                CompareKeyed(left.ThreadedComments.Select(comment => $"{comment.CellReference}|{comment.Author}|{comment.Text}|{comment.Done}"), right.ThreadedComments.Select(comment => $"{comment.CellReference}|{comment.Author}|{comment.Text}|{comment.Done}"), "ThreadedComment", sheetName, differences, maxDifferences);
            }
        }

        private void CompareNamedRanges(ExcelDocument other, ICollection<ExcelWorkbookDifference> differences, int maxDifferences) {
            var left = ListNamedRanges(includeBuiltIn: true, includeHidden: true).Select(name => $"{name.SheetName}|{name.Name}|{name.Reference}|{name.Hidden}").OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToArray();
            var right = other.ListNamedRanges(includeBuiltIn: true, includeHidden: true).Select(name => $"{name.SheetName}|{name.Name}|{name.Reference}|{name.Hidden}").OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToArray();
            CompareKeyed(left, right, "NamedRange", null, differences, maxDifferences);
        }

        private static void CompareKeyed(IEnumerable<string> leftItems, IEnumerable<string> rightItems, string category, string? sheetName, ICollection<ExcelWorkbookDifference> differences, int maxDifferences) {
            var left = new HashSet<string>(leftItems, StringComparer.OrdinalIgnoreCase);
            var right = new HashSet<string>(rightItems, StringComparer.OrdinalIgnoreCase);
            foreach (string item in left.Where(item => !right.Contains(item))) {
                AddDifference(differences, maxDifferences, new ExcelWorkbookDifference(category, $"{category} exists only in left workbook.", sheetName, leftValue: item));
                if (differences.Count >= maxDifferences) return;
            }
            foreach (string item in right.Where(item => !left.Contains(item))) {
                AddDifference(differences, maxDifferences, new ExcelWorkbookDifference(category, $"{category} exists only in right workbook.", sheetName, rightValue: item));
                if (differences.Count >= maxDifferences) return;
            }
        }

        private static void CompareScalar(string category, string message, string sheetName, string? leftValue, string? rightValue, ICollection<ExcelWorkbookDifference> differences, int maxDifferences) {
            if (!string.Equals(leftValue ?? string.Empty, rightValue ?? string.Empty, StringComparison.Ordinal)) {
                AddDifference(differences, maxDifferences, new ExcelWorkbookDifference(category, message, sheetName, leftValue: leftValue, rightValue: rightValue));
            }
        }

        private uint GetNextPowerQueryConnectionId() {
            uint maxId = 0;
            foreach (OpenXmlPart part in EnumerateWorkbookConnectionParts()) {
                try {
                    XDocument document = XDocument.Parse(ReadOpenXmlPartText(part));
                    foreach (XElement connection in document.Descendants().Where(element => element.Name.LocalName == "connection")) {
                        if (uint.TryParse(connection.Attribute("id")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint id)) {
                            maxId = Math.Max(maxId, id);
                        }
                    }
                } catch {
                    // Extended metadata can be caller-supplied and not all parts are XML connection parts.
                }
            }

            maxId = Math.Max(maxId, (uint)InspectDataModel().ConnectionPartCount);
            return maxId + 1U;
        }

        private static string BuildConnectionMetadataXml(string name, string? description, string commandText, bool refreshOnOpen, uint connectionId) {
            XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var connection = new XElement(main + "connection",
                new XAttribute("id", connectionId.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("name", name),
                new XAttribute("type", "5"),
                new XAttribute("refreshedVersion", "7"),
                new XAttribute("refreshOnLoad", refreshOnOpen ? "1" : "0"));
            if (!string.IsNullOrWhiteSpace(description)) {
                connection.SetAttributeValue("description", description!.Trim());
            }
            if (!string.IsNullOrWhiteSpace(commandText)) {
                connection.Add(new XElement(main + "dbPr",
                    new XAttribute("connection", "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" + name + ";Extended Properties=\"\""),
                    new XAttribute("command", commandText),
                    new XAttribute("commandType", "8")));
            }

            return new XDocument(new XElement(main + "connections", new XAttribute("count", "1"), connection)).ToString(SaveOptions.DisableFormatting);
        }

        private static string NormalizeQueryTableName(string? queryTableName, string connectionName) {
            return string.IsNullOrWhiteSpace(queryTableName) ? connectionName + "Table" : queryTableName!.Trim();
        }

        private static string BuildQueryTableMetadataXml(string queryTableName, uint connectionId) {
            XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            return new XDocument(new XElement(main + "queryTable",
                new XAttribute("name", queryTableName),
                new XAttribute("connectionId", connectionId.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("autoFormatId", "16"),
                new XAttribute("applyNumberFormats", "0"),
                new XAttribute("applyBorderFormats", "0"),
                new XAttribute("applyFontFormats", "0"),
                new XAttribute("applyPatternFormats", "0"),
                new XAttribute("applyAlignmentFormats", "0"),
                new XAttribute("applyWidthHeightFormats", "0"))).ToString(SaveOptions.DisableFormatting);
        }
    }
}
