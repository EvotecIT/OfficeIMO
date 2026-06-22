using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

#pragma warning disable CS1591
namespace OfficeIMO.Excel {
    /// <summary>
    /// Severity assigned to workbook diagnostics and compliance findings.
    /// </summary>
    public enum ExcelFindingSeverity {
        Info,
        Warning,
        Error
    }

    /// <summary>
    /// A diagnostic issue discovered in an Excel workbook.
    /// </summary>
    public sealed class ExcelWorkbookDiagnosticIssue {
        internal ExcelWorkbookDiagnosticIssue(string category, ExcelFindingSeverity severity, string message, string? sheetName = null, string? address = null, string? repairAction = null) {
            Category = category;
            Severity = severity;
            Message = message;
            SheetName = sheetName;
            Address = address;
            RepairAction = repairAction;
        }

        public string Category { get; }
        public ExcelFindingSeverity Severity { get; }
        public string Message { get; }
        public string? SheetName { get; }
        public string? Address { get; }
        public string? RepairAction { get; }
    }

    /// <summary>
    /// Diagnostic report produced by the workbook doctor.
    /// </summary>
    public sealed class ExcelWorkbookDiagnosticReport {
        internal ExcelWorkbookDiagnosticReport(IReadOnlyList<ExcelWorkbookDiagnosticIssue> issues, int repairedIssueCount) {
            Issues = issues;
            RepairedIssueCount = repairedIssueCount;
        }

        public IReadOnlyList<ExcelWorkbookDiagnosticIssue> Issues { get; }
        public int RepairedIssueCount { get; }
        public bool HasErrors => Issues.Any(issue => issue.Severity == ExcelFindingSeverity.Error);
        public bool HasWarnings => Issues.Any(issue => issue.Severity == ExcelFindingSeverity.Warning);
    }

    /// <summary>
    /// Options for workbook diagnostic scans and safe repairs.
    /// </summary>
    public sealed class ExcelWorkbookDoctorOptions {
        public bool ValidateOpenXml { get; set; } = true;
        public bool CheckDefinedNames { get; set; } = true;
        public bool CheckFormulas { get; set; } = true;
        public bool CheckTables { get; set; } = true;
        public bool CheckDrawings { get; set; } = true;
        public bool CheckConnections { get; set; } = true;
        public bool RepairDefinedNames { get; set; }
    }

    /// <summary>
    /// Formula metadata discovered during workbook analysis.
    /// </summary>
    public sealed class ExcelFormulaInfo {
        internal ExcelFormulaInfo(string sheetName, string address, string formula, IReadOnlyList<string> references, IReadOnlyList<string> functions, bool hasExternalReference, bool isVolatile) {
            SheetName = sheetName;
            Address = address;
            Formula = formula;
            References = references;
            Functions = functions;
            HasExternalReference = hasExternalReference;
            IsVolatile = isVolatile;
        }

        public string SheetName { get; }
        public string Address { get; }
        public string Formula { get; }
        public IReadOnlyList<string> References { get; }
        public IReadOnlyList<string> Functions { get; }
        public bool HasExternalReference { get; }
        public bool IsVolatile { get; }
    }

    /// <summary>
    /// Workbook formula intelligence report.
    /// </summary>
    public sealed class ExcelFormulaAnalysisReport {
        internal ExcelFormulaAnalysisReport(IReadOnlyList<ExcelFormulaInfo> formulas) {
            Formulas = formulas;
        }

        public IReadOnlyList<ExcelFormulaInfo> Formulas { get; }
        public int FormulaCount => Formulas.Count;
        public int VolatileFormulaCount => Formulas.Count(formula => formula.IsVolatile);
        public int ExternalReferenceCount => Formulas.Count(formula => formula.HasExternalReference);
    }

    /// <summary>
    /// Named range metadata with scope and hidden state.
    /// </summary>
    public sealed class ExcelNamedRangeInfo {
        internal ExcelNamedRangeInfo(string name, string reference, string? sheetName, bool hidden, bool builtIn) {
            Name = name;
            Reference = reference;
            SheetName = sheetName;
            Hidden = hidden;
            BuiltIn = builtIn;
        }

        public string Name { get; }
        public string Reference { get; }
        public string? SheetName { get; }
        public bool Hidden { get; }
        public bool BuiltIn { get; }
    }

    /// <summary>
    /// Result of importing normalized delimited text.
    /// </summary>
    public sealed class ExcelDelimitedImportResult {
        internal ExcelDelimitedImportResult(ExcelDataSetImportResult importResult, char delimiter, string encodingName, IReadOnlyList<string> warnings) {
            ImportResult = importResult;
            Delimiter = delimiter;
            EncodingName = encodingName;
            Warnings = warnings;
        }

        public ExcelDataSetImportResult ImportResult { get; }
        public char Delimiter { get; }
        public string EncodingName { get; }
        public IReadOnlyList<string> Warnings { get; }
    }

    /// <summary>
    /// Options for culture-aware CSV/TSV import normalization.
    /// </summary>
    public sealed class ExcelDelimitedImportOptions {
        public char? Delimiter { get; set; }
        public bool HeadersInFirstRow { get; set; } = true;
        public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
        public bool ConvertNumbersAndDates { get; set; } = true;
        public bool CreateTable { get; set; } = true;
        public string? SheetName { get; set; }
        public string? TableName { get; set; }
        public TableStyle TableStyle { get; set; } = TableStyle.TableStyleMedium2;
    }

    /// <summary>
    /// One workbook difference discovered by a structural/value comparison.
    /// </summary>
    public sealed class ExcelWorkbookDifference {
        internal ExcelWorkbookDifference(string category, string message, string? sheetName = null, string? address = null, string? leftValue = null, string? rightValue = null) {
            Category = category;
            Message = message;
            SheetName = sheetName;
            Address = address;
            LeftValue = leftValue;
            RightValue = rightValue;
        }

        public string Category { get; }
        public string Message { get; }
        public string? SheetName { get; }
        public string? Address { get; }
        public string? LeftValue { get; }
        public string? RightValue { get; }
    }

    /// <summary>
    /// Workbook comparison report.
    /// </summary>
    public sealed class ExcelWorkbookDiffReport {
        internal ExcelWorkbookDiffReport(IReadOnlyList<ExcelWorkbookDifference> differences) {
            Differences = differences;
        }

        public IReadOnlyList<ExcelWorkbookDifference> Differences { get; }
        public bool AreEqual => Differences.Count == 0;
    }

    /// <summary>
    /// Accessibility and compliance finding.
    /// </summary>
    public sealed class ExcelAccessibilityFinding {
        internal ExcelAccessibilityFinding(string category, ExcelFindingSeverity severity, string message, string? sheetName = null, string? address = null) {
            Category = category;
            Severity = severity;
            Message = message;
            SheetName = sheetName;
            Address = address;
        }

        public string Category { get; }
        public ExcelFindingSeverity Severity { get; }
        public string Message { get; }
        public string? SheetName { get; }
        public string? Address { get; }
    }

    /// <summary>
    /// Workbook accessibility and compliance report.
    /// </summary>
    public sealed class ExcelAccessibilityReport {
        internal ExcelAccessibilityReport(IReadOnlyList<ExcelAccessibilityFinding> findings) {
            Findings = findings;
        }

        public IReadOnlyList<ExcelAccessibilityFinding> Findings { get; }
        public bool HasWarnings => Findings.Any(finding => finding.Severity != ExcelFindingSeverity.Info);
    }

    /// <summary>
    /// Explicit large-workbook streaming capability contract for the current workbook.
    /// </summary>
    public sealed class ExcelStreamingContractReport {
        internal ExcelStreamingContractReport(int worksheetCount, int estimatedCellCount, bool hasDirectDataSetFastSaveState, bool hasDeferredDirectDataSetImport, string recommendation) {
            WorksheetCount = worksheetCount;
            EstimatedCellCount = estimatedCellCount;
            HasDirectDataSetFastSaveState = hasDirectDataSetFastSaveState;
            HasDeferredDirectDataSetImport = hasDeferredDirectDataSetImport;
            Recommendation = recommendation;
        }

        public int WorksheetCount { get; }
        public int EstimatedCellCount { get; }
        public bool HasDirectDataSetFastSaveState { get; }
        public bool HasDeferredDirectDataSetImport { get; }
        public string Recommendation { get; }
    }

    /// <summary>
    /// Data-model and external-query awareness report.
    /// </summary>
    public sealed class ExcelDataModelReport {
        internal ExcelDataModelReport(int connectionPartCount, int queryTablePartCount, int modelPartCount, int externalLinkPartCount, IReadOnlyList<string> details) {
            ConnectionPartCount = connectionPartCount;
            QueryTablePartCount = queryTablePartCount;
            ModelPartCount = modelPartCount;
            ExternalLinkPartCount = externalLinkPartCount;
            Details = details;
        }

        public int ConnectionPartCount { get; }
        public int QueryTablePartCount { get; }
        public int ModelPartCount { get; }
        public int ExternalLinkPartCount { get; }
        public IReadOnlyList<string> Details { get; }
        public bool HasDataModelOrQueries => ConnectionPartCount > 0 || QueryTablePartCount > 0 || ModelPartCount > 0 || ExternalLinkPartCount > 0;
    }

    public partial class ExcelDocument {
        private static readonly Regex FormulaReferenceRegex = new Regex(@"(?<![A-Za-z0-9_])(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)?!?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?", RegexOptions.Compiled);
        private static readonly Regex FormulaFunctionRegex = new Regex(@"(?<![A-Za-z0-9_])([A-Za-z_][A-Za-z0-9_\.]*)\s*\(", RegexOptions.Compiled);
        private static readonly HashSet<string> VolatileFormulaFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"
        };

        /// <summary>
        /// Runs workbook diagnostics and optional safe repairs for common automation hazards.
        /// </summary>
        public ExcelWorkbookDiagnosticReport RunWorkbookDoctor(ExcelWorkbookDoctorOptions? options = null) {
            options ??= new ExcelWorkbookDoctorOptions();
            int repaired = 0;
            if (options.RepairDefinedNames) {
                int before = WorkbookRoot.DefinedNames?.Elements<DefinedName>().Count() ?? 0;
                RepairDefinedNames(save: true);
                int after = WorkbookRoot.DefinedNames?.Elements<DefinedName>().Count() ?? 0;
                repaired += Math.Max(0, before - after);
            }

            var issues = new List<ExcelWorkbookDiagnosticIssue>();
            if (options.ValidateOpenXml) {
                foreach (ValidationErrorInfo error in new OpenXmlValidator().Validate(_spreadSheetDocument).Take(50)) {
                    issues.Add(new ExcelWorkbookDiagnosticIssue("OpenXml", ExcelFindingSeverity.Error, error.Description ?? "OpenXml validation error.", repairAction: "Inspect package XML or save with validation enabled."));
                }
            }

            if (options.CheckDefinedNames) AddDefinedNameDiagnostics(issues);
            if (options.CheckFormulas) AddFormulaDiagnostics(issues);
            if (options.CheckTables) AddTableDiagnostics(issues);
            if (options.CheckDrawings) AddDrawingDiagnostics(issues);
            if (options.CheckConnections) AddConnectionDiagnostics(issues);

            return new ExcelWorkbookDiagnosticReport(issues, repaired);
        }

        /// <summary>
        /// Analyzes formulas for references, external links, and volatile functions.
        /// </summary>
        public ExcelFormulaAnalysisReport AnalyzeFormulas() {
            var formulas = new List<ExcelFormulaInfo>();
            foreach (var item in EnumerateWorksheetParts()) {
                Worksheet? worksheet = item.WorksheetPart.Worksheet;
                if (worksheet == null) continue;

                foreach (Cell cell in worksheet.Descendants<Cell>()) {
                    string? formulaText = cell.CellFormula?.Text;
                    if (string.IsNullOrWhiteSpace(formulaText)) continue;

                    string address = cell.CellReference?.Value ?? string.Empty;
                    var references = FormulaReferenceRegex.Matches(formulaText!)
                        .Cast<Match>()
                        .Select(match => match.Value)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToArray();
                    var functions = FormulaFunctionRegex.Matches(formulaText!)
                        .Cast<Match>()
                        .Select(match => match.Groups[1].Value)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToArray();

                    formulas.Add(new ExcelFormulaInfo(
                        item.SheetName,
                        address,
                        formulaText!,
                        references,
                        functions,
                        formulaText!.IndexOf('[') >= 0,
                        functions.Any(function => VolatileFormulaFunctions.Contains(function))));
                }
            }

            return new ExcelFormulaAnalysisReport(formulas);
        }

        /// <summary>
        /// Returns all workbook and sheet-scoped defined names with scope and hidden metadata.
        /// </summary>
        public IReadOnlyList<ExcelNamedRangeInfo> ListNamedRanges(bool includeBuiltIn = true, bool includeHidden = true) {
            var definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) return Array.Empty<ExcelNamedRangeInfo>();

            var sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            var result = new List<ExcelNamedRangeInfo>();
            foreach (DefinedName name in definedNames.Elements<DefinedName>()) {
                string nameValue = name.Name?.Value ?? string.Empty;
                bool builtIn = nameValue.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);
                bool hidden = name.Hidden?.Value ?? false;
                if ((!includeBuiltIn && builtIn) || (!includeHidden && hidden)) continue;

                string? sheetName = null;
                if (name.LocalSheetId?.Value is uint local && local < sheets.Count) {
                    sheetName = sheets[(int)local].Name?.Value;
                }

                result.Add(new ExcelNamedRangeInfo(nameValue, name.Text ?? string.Empty, sheetName, hidden, builtIn));
            }

            return result;
        }

        /// <summary>
        /// Renames an existing workbook or sheet-scoped defined name.
        /// </summary>
        public bool RenameNamedRange(string oldName, string newName, ExcelSheet? scope = null, NameValidationMode validationMode = NameValidationMode.Sanitize, bool save = true) {
            if (string.IsNullOrWhiteSpace(oldName)) throw new ArgumentException("Old name is required.", nameof(oldName));
            if (string.IsNullOrWhiteSpace(newName)) throw new ArgumentException("New name is required.", nameof(newName));

            string? reference = GetNamedRange(oldName, scope);
            if (reference == null) return false;

            bool hidden = ListNamedRanges(includeBuiltIn: true, includeHidden: true)
                .FirstOrDefault(item => string.Equals(item.Name, oldName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(item.SheetName, scope?.Name, StringComparison.OrdinalIgnoreCase))?.Hidden ?? false;

            RemoveNamedRange(oldName, scope, save: false);
            SetNamedRange(newName, reference, scope, save: save, hidden: hidden, validationMode: validationMode);
            return true;
        }

        /// <summary>
        /// Imports normalized CSV/TSV content into a worksheet using OfficeIMO's tabular writer.
        /// </summary>
        public ExcelDelimitedImportResult ImportDelimitedText(string text, ExcelDelimitedImportOptions? options = null) {
            if (text == null) throw new ArgumentNullException(nameof(text));
            options ??= new ExcelDelimitedImportOptions();
            char delimiter = options.Delimiter ?? DetectDelimiter(text);
            var warnings = new List<string>();
            DataTable table = ParseDelimitedText(text, delimiter, options, warnings);
            table.TableName = string.IsNullOrWhiteSpace(options.SheetName) ? "Import" : options.SheetName!.Trim();

            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            IReadOnlyList<ExcelDataSetImportResult> results = InsertDataSet(dataSet, createTables: options.CreateTable, tableStyle: options.TableStyle, includeHeaders: true, includeAutoFilter: true, autoFit: false);
            return new ExcelDelimitedImportResult(results[0], delimiter, "UTF-16 string", warnings);
        }

        /// <summary>
        /// Compares this workbook with another workbook by sheets, dimensions, formulas, and visible cell text.
        /// </summary>
        public ExcelWorkbookDiffReport CompareWorkbook(ExcelDocument other, int maxDifferences = 200) {
            if (other == null) throw new ArgumentNullException(nameof(other));
            var differences = new List<ExcelWorkbookDifference>();
            var otherSheets = other.Sheets.ToDictionary(sheet => sheet.Name, StringComparer.OrdinalIgnoreCase);

            foreach (ExcelSheet sheet in Sheets) {
                if (!otherSheets.TryGetValue(sheet.Name, out ExcelSheet? otherSheet)) {
                    AddDifference(differences, maxDifferences, new ExcelWorkbookDifference("Sheet", "Sheet exists only in left workbook.", sheet.Name));
                    continue;
                }

                CompareSheetValues(sheet, otherSheet, differences, maxDifferences);
                if (differences.Count >= maxDifferences) break;
            }

            var leftNames = new HashSet<string>(Sheets.Select(sheet => sheet.Name), StringComparer.OrdinalIgnoreCase);
            foreach (ExcelSheet sheet in other.Sheets.Where(sheet => !leftNames.Contains(sheet.Name))) {
                AddDifference(differences, maxDifferences, new ExcelWorkbookDifference("Sheet", "Sheet exists only in right workbook.", sheet.Name));
            }

            return new ExcelWorkbookDiffReport(differences);
        }

        /// <summary>
        /// Reports workbook accessibility and compliance issues such as missing alt text, hidden sheets, merged cells, and missing table headers.
        /// </summary>
        public ExcelAccessibilityReport AnalyzeAccessibility() {
            var findings = new List<ExcelAccessibilityFinding>();
            foreach (ExcelSheet sheet in Sheets) {
                foreach (ExcelImage image in sheet.Images) {
                    if (string.IsNullOrWhiteSpace(image.Description) && string.IsNullOrWhiteSpace(image.Title)) {
                        findings.Add(new ExcelAccessibilityFinding("ImageAltText", ExcelFindingSeverity.Warning, "Image has no title or alternative text.", sheet.Name, A1.CellReference(image.RowIndex, image.ColumnIndex)));
                    }
                }
            }

            ExcelWorkbookSnapshot snapshot = CreateInspectionSnapshot();
            foreach (ExcelWorksheetSnapshot worksheet in snapshot.Worksheets) {
                if (worksheet.Hidden) {
                    findings.Add(new ExcelAccessibilityFinding("HiddenData", ExcelFindingSeverity.Info, "Worksheet is hidden and may contain data excluded from normal review.", worksheet.Name));
                }

                foreach (ExcelMergedRangeSnapshot merge in worksheet.MergedRanges) {
                    findings.Add(new ExcelAccessibilityFinding("MergedCells", ExcelFindingSeverity.Warning, "Merged cells can make sorting, filtering, and screen-reader navigation harder.", worksheet.Name, merge.A1Range));
                }

                foreach (ExcelTableSnapshot table in worksheet.Tables.Where(table => !table.HasHeaderRow)) {
                    findings.Add(new ExcelAccessibilityFinding("TableHeaders", ExcelFindingSeverity.Warning, $"Table '{table.Name}' has no header row.", worksheet.Name, table.A1Range));
                }
            }

            return new ExcelAccessibilityReport(findings);
        }

        /// <summary>
        /// Returns a streaming/read-write contract summary for large workbook workflows.
        /// </summary>
        public ExcelStreamingContractReport GetStreamingContract() {
            ExcelWorkbookSnapshot snapshot = CreateInspectionSnapshot();
            int estimatedCellCount = snapshot.Worksheets.Sum(sheet => sheet.Cells.Count);
            string recommendation = HasDeferredDirectDataSetImport || HasDirectDataSetFastSaveState
                ? "Workbook is currently using OfficeIMO direct tabular save state."
                : estimatedCellCount > 500000
                    ? "Use direct DataSet/DataTable writers or reader streaming APIs for this workbook size."
                    : "Standard OfficeIMO workbook APIs are suitable for this workbook size.";
            return new ExcelStreamingContractReport(snapshot.Worksheets.Count, estimatedCellCount, HasDirectDataSetFastSaveState, HasDeferredDirectDataSetImport, recommendation);
        }

        /// <summary>
        /// Inspects data model, Power Query, connection, query table, and external link package parts.
        /// </summary>
        public ExcelDataModelReport InspectDataModel() {
            var details = new List<string>();
            int connection = 0;
            int queryTable = 0;
            int model = 0;
            int external = 0;
            foreach (OpenXmlPart part in EnumerateParts(WorkbookPartRoot)) {
                string contentType = part.ContentType ?? string.Empty;
                string uri = part.Uri.ToString();
                if (contentType.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0) { connection++; details.Add(uri); }
                if (contentType.IndexOf("queryTable", StringComparison.OrdinalIgnoreCase) >= 0) { queryTable++; details.Add(uri); }
                if (contentType.IndexOf("model", StringComparison.OrdinalIgnoreCase) >= 0) { model++; details.Add(uri); }
                if (contentType.IndexOf("externalLink", StringComparison.OrdinalIgnoreCase) >= 0) { external++; details.Add(uri); }
            }

            return new ExcelDataModelReport(connection, queryTable, model, external, details.Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
        }

        private void AddDefinedNameDiagnostics(ICollection<ExcelWorkbookDiagnosticIssue> issues) {
            var names = ListNamedRanges(includeBuiltIn: true, includeHidden: true);
            foreach (var duplicate in names.GroupBy(name => (name.SheetName ?? string.Empty) + "|" + name.Name, StringComparer.OrdinalIgnoreCase).Where(group => group.Count() > 1)) {
                issues.Add(new ExcelWorkbookDiagnosticIssue("DefinedName", ExcelFindingSeverity.Warning, $"Duplicate defined name '{duplicate.First().Name}' in the same scope.", duplicate.First().SheetName, repairAction: "Run workbook doctor with RepairDefinedNames enabled."));
            }

            foreach (ExcelNamedRangeInfo name in names.Where(name => name.Reference.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0)) {
                issues.Add(new ExcelWorkbookDiagnosticIssue("DefinedName", ExcelFindingSeverity.Error, $"Defined name '{name.Name}' contains #REF!.", name.SheetName, repairAction: "Remove or recreate the defined name."));
            }
        }

        private void AddFormulaDiagnostics(ICollection<ExcelWorkbookDiagnosticIssue> issues) {
            foreach (ExcelFormulaInfo formula in AnalyzeFormulas().Formulas) {
                if (formula.HasExternalReference) {
                    issues.Add(new ExcelWorkbookDiagnosticIssue("Formula", ExcelFindingSeverity.Warning, "Formula references an external workbook.", formula.SheetName, formula.Address, "Review external links before automated refresh or distribution."));
                }
                if (formula.IsVolatile) {
                    issues.Add(new ExcelWorkbookDiagnosticIssue("Formula", ExcelFindingSeverity.Info, "Formula uses a volatile function.", formula.SheetName, formula.Address));
                }
            }
        }

        private void AddTableDiagnostics(ICollection<ExcelWorkbookDiagnosticIssue> issues) {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in EnumerateWorksheetParts()) {
                foreach (TableDefinitionPart tablePart in item.WorksheetPart.TableDefinitionParts) {
                    string name = tablePart.Table?.Name?.Value ?? tablePart.Table?.DisplayName?.Value ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(name) && !names.Add(name)) {
                        issues.Add(new ExcelWorkbookDiagnosticIssue("Table", ExcelFindingSeverity.Error, $"Duplicate table name '{name}'.", item.SheetName, repairAction: "Rename one table before saving."));
                    }
                }
            }
        }

        private void AddDrawingDiagnostics(ICollection<ExcelWorkbookDiagnosticIssue> issues) {
            foreach (ExcelSheet sheet in Sheets) {
                foreach (ExcelImage image in sheet.Images) {
                    if (image.GetBytes().Length == 0) {
                        issues.Add(new ExcelWorkbookDiagnosticIssue("Drawing", ExcelFindingSeverity.Warning, $"Image '{image.Name}' has no readable image bytes.", sheet.Name, A1.CellReference(image.RowIndex, image.ColumnIndex)));
                    }
                }
            }
        }

        private void AddConnectionDiagnostics(ICollection<ExcelWorkbookDiagnosticIssue> issues) {
            ExcelDataModelReport dataModel = InspectDataModel();
            if (dataModel.HasDataModelOrQueries) {
                issues.Add(new ExcelWorkbookDiagnosticIssue("DataModel", ExcelFindingSeverity.Info, "Workbook contains connection, query, data-model, or external-link parts that OfficeIMO preserves but does not execute.", repairAction: "Use refresh-on-open metadata or Excel to execute queries."));
            }
        }

        private IEnumerable<(string SheetName, WorksheetPart WorksheetPart)> EnumerateWorksheetParts() {
            WorkbookPart workbookPart = WorkbookPartRoot;
            foreach (Sheet sheet in workbookPart.Workbook?.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheet.Id?.Value == null) continue;
                if (workbookPart.GetPartById(sheet.Id.Value) is WorksheetPart worksheetPart) {
                    yield return (sheet.Name?.Value ?? string.Empty, worksheetPart);
                }
            }
        }

        private static IEnumerable<OpenXmlPart> EnumerateParts(OpenXmlPartContainer container) {
            foreach (IdPartPair pair in container.Parts) {
                yield return pair.OpenXmlPart;
                foreach (OpenXmlPart child in EnumerateParts(pair.OpenXmlPart)) {
                    yield return child;
                }
            }
        }

        private static char DetectDelimiter(string text) {
            string firstLine = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).FirstOrDefault() ?? string.Empty;
            var candidates = new[] { ',', ';', '\t', '|' };
            return candidates
                .Select(candidate => new { Delimiter = candidate, Count = firstLine.Count(ch => ch == candidate) })
                .OrderByDescending(item => item.Count)
                .First().Delimiter;
        }

        private static DataTable ParseDelimitedText(string text, char delimiter, ExcelDelimitedImportOptions options, ICollection<string> warnings) {
            var rows = ParseDelimitedRows(text, delimiter).ToList();
            var table = new DataTable { Locale = options.Culture };
            if (rows.Count == 0) return table;

            int columnCount = rows.Max(row => row.Count);
            int dataStart = options.HeadersInFirstRow ? 1 : 0;
            for (int i = 0; i < columnCount; i++) {
                string name = options.HeadersInFirstRow && i < rows[0].Count && !string.IsNullOrWhiteSpace(rows[0][i])
                    ? rows[0][i]
                    : "Column" + (i + 1).ToString(CultureInfo.InvariantCulture);
                if (table.Columns.Contains(name)) name += "_" + (i + 1).ToString(CultureInfo.InvariantCulture);
                table.Columns.Add(name, typeof(object));
            }

            for (int rowIndex = dataStart; rowIndex < rows.Count; rowIndex++) {
                DataRow row = table.NewRow();
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                    string value = columnIndex < rows[rowIndex].Count ? rows[rowIndex][columnIndex] : string.Empty;
                    row[columnIndex] = ConvertDelimitedValue(value, options, warnings);
                }
                table.Rows.Add(row);
            }

            return table;
        }

        private static object ConvertDelimitedValue(string value, ExcelDelimitedImportOptions options, ICollection<string> warnings) {
            if (value.Length == 0) return DBNull.Value;
            if (!options.ConvertNumbersAndDates) return value;
            if (decimal.TryParse(value, NumberStyles.Number, options.Culture, out decimal number)) return number;
            if (DateTime.TryParse(value, options.Culture, DateTimeStyles.None, out DateTime date)) return date;
            return value;
        }

        private static IEnumerable<List<string>> ParseDelimitedRows(string text, char delimiter) {
            var row = new List<string>();
            var field = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (quoted) {
                    if (ch == '"' && i + 1 < text.Length && text[i + 1] == '"') {
                        field.Append('"');
                        i++;
                    } else if (ch == '"') {
                        quoted = false;
                    } else {
                        field.Append(ch);
                    }
                    continue;
                }

                if (ch == '"') { quoted = true; continue; }
                if (ch == delimiter) { row.Add(field.ToString()); field.Clear(); continue; }
                if (ch == '\r') continue;
                if (ch == '\n') {
                    row.Add(field.ToString());
                    field.Clear();
                    yield return row;
                    row = new List<string>();
                    continue;
                }

                field.Append(ch);
            }

            row.Add(field.ToString());
            if (row.Count > 1 || row[0].Length > 0) yield return row;
        }

        private static void CompareSheetValues(ExcelSheet left, ExcelSheet right, ICollection<ExcelWorkbookDifference> differences, int maxDifferences) {
            Worksheet? leftWorksheet = left.WorksheetPart.Worksheet;
            Worksheet? rightWorksheet = right.WorksheetPart.Worksheet;
            string leftRange = leftWorksheet?.SheetDimension?.Reference?.Value ?? string.Empty;
            string rightRange = rightWorksheet?.SheetDimension?.Reference?.Value ?? string.Empty;
            if (!string.Equals(leftRange, rightRange, StringComparison.OrdinalIgnoreCase)) {
                AddDifference(differences, maxDifferences, new ExcelWorkbookDifference("UsedRange", "Used ranges differ.", left.Name, leftValue: leftRange, rightValue: rightRange));
            }

            var rightCells = (rightWorksheet?.Descendants<Cell>() ?? Enumerable.Empty<Cell>())
                .Where(cell => !string.IsNullOrWhiteSpace(cell.CellReference?.Value))
                .ToDictionary(cell => cell.CellReference!.Value!, StringComparer.OrdinalIgnoreCase);

            foreach (Cell leftCell in leftWorksheet?.Descendants<Cell>() ?? Enumerable.Empty<Cell>()) {
                if (differences.Count >= maxDifferences) break;
                string? address = leftCell.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(address)) continue;

                rightCells.TryGetValue(address!, out Cell? rightCell);
                string leftValue = left.GetCellText(leftCell);
                string rightValue = rightCell == null ? string.Empty : right.GetCellText(rightCell);
                string? leftFormula = leftCell.CellFormula?.Text;
                string? rightFormula = rightCell?.CellFormula?.Text;
                if (!string.Equals(leftValue, rightValue, StringComparison.Ordinal) || !string.Equals(leftFormula, rightFormula, StringComparison.Ordinal)) {
                    AddDifference(differences, maxDifferences, new ExcelWorkbookDifference("Cell", "Cell value or formula differs.", left.Name, address, leftFormula ?? leftValue, rightFormula ?? rightValue));
                }
            }
        }

        private static void AddDifference(ICollection<ExcelWorkbookDifference> differences, int maxDifferences, ExcelWorkbookDifference difference) {
            if (differences.Count < maxDifferences) differences.Add(difference);
        }
    }
}
#pragma warning restore CS1591
