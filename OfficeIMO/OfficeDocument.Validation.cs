using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO;

public static partial class OfficeDocument {
    /// <summary>Validates a file through its native engine, opening it read-only and never saving it.</summary>
    /// <remarks>Unsupported checks produce explicit not-applicable diagnostics. Legacy Office loaders validate their projected model, not a native binary schema.</remarks>
    public static OfficeValidationReport Validate(string path, OfficeDocumentCheck[]? checks = null,
        OfficeDocumentValidationOptions? options = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A document path is required.", nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException("Document not found.", path);
        var settings = options ?? new OfficeDocumentValidationOptions();
        var selected = ValidationChecks(checks, settings);
        var diagnostics = new List<OfficeDiagnostic>();
        var evidence = new List<object>();
        cancellationToken.ThrowIfCancellationRequested();
        string extension = Path.GetExtension(path).ToLowerInvariant();
        diagnostics.Add(new OfficeDiagnostic("Validation.Source", OfficeDiagnosticSeverity.Information,
            "Source document.", Path.GetFullPath(path), extension));
        bool package = extension != ".pdf";
        if (selected.Contains(OfficeDocumentCheck.Security) && package) {
            using var source = File.OpenRead(path);
            OfficePackageSecurityReport report;
            try { report = OfficePackageSecurityInspector.Inspect(source, ValidationSecurity(settings)); }
            catch (OfficePackageSecurityException error) {
                // Streaming admission stops before parsing; retain the typed rule, observed value and limit as native evidence.
                evidence.Add(error);
                diagnostics.Add(new OfficeDiagnostic("Security." + error.Rule, OfficeDiagnosticSeverity.Error, error.Message, error.PartName, extension));
                if (selected.Any(check => check != OfficeDocumentCheck.Security)) diagnostics.Add(new OfficeDiagnostic("Validation.CheckSkipped",
                    OfficeDiagnosticSeverity.Warning, "Remaining checks were skipped because package security rejected the input.", format: extension));
                return new OfficeValidationReport(diagnostics, evidence);
            }
            AddValidationSecurity(report, diagnostics, evidence);
            if (report.ContainerKind == OfficePackageContainerKind.Unknown)
                ValidationNotApplicable(diagnostics, OfficeDocumentCheck.Security, extension, "No supported package container was found.");
            selected = selected.Where(check => check != OfficeDocumentCheck.Security).ToArray();
            // Do not pass an input rejected by the security admission policy to a format parser.
            if (!report.IsValid) {
                if (selected.Length != 0) diagnostics.Add(new OfficeDiagnostic("Validation.CheckSkipped", OfficeDiagnosticSeverity.Warning,
                    "Remaining checks were skipped because package security rejected the input.", format: extension));
                return new OfficeValidationReport(diagnostics, evidence);
            }
        }
        if (selected.Length == 0) return new OfficeValidationReport(diagnostics, evidence);
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions? pdfRead = extension == ".pdf" ? ValidationPdfRead(settings) : null;
        // Every owned Office model uses explicit persistence. Borrowed models use the object overload below.
        object? document = extension switch {
            ".docx" or ".docm" or ".dotx" or ".dotm" or ".doc" or ".dot" => WordDocument.Load(path,
                new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit, Password = settings.Password }),
            ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xls" or ".xlt" or ".xlsb" => ExcelDocument.Load(path,
                new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit, Password = settings.Password }),
            ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" or ".ppt" or ".pps" or ".pot" => PowerPointDocument.Load(path,
                new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit, Password = settings.Password }),
            ".pdf" => PdfDocument.Load(path, pdfRead, cancellationToken),
            _ => null
        };
        if (document == null) {
            foreach (var check in selected) ValidationNotApplicable(diagnostics, check, extension, "No native validator is registered in OfficeIMO for this format.");
        } else {
            try {
                ValidateModel(document, selected, settings, diagnostics, evidence, cancellationToken, path, pdfRead);
                if (extension is ".doc" or ".dot" or ".xls" or ".xlt" or ".xlsb" or ".ppt" or ".pps" or ".pot")
                    diagnostics.Add(new OfficeDiagnostic("Validation.ProjectedModel", OfficeDiagnosticSeverity.Information,
                        "Structure and schema checks describe the loaded Office model; they do not certify the source binary format.", format: extension));
            } finally { (document as IDisposable)?.Dispose(); }
        }
        return new OfficeValidationReport(diagnostics, evidence);
    }

    /// <summary>Validates the current native model without saving, reopening its path, repairing it, or disposing caller-owned state.</summary>
    public static OfficeValidationReport Validate(object document, OfficeDocumentCheck[]? checks = null,
        OfficeDocumentValidationOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document is string path) return Validate(path, checks, options, cancellationToken);
        var settings = options ?? new OfficeDocumentValidationOptions();
        if (settings.Password != null) throw new ArgumentException("Password applies only to path validation.", nameof(options));
        var selected = ValidationChecks(checks, settings);
        var diagnostics = new List<OfficeDiagnostic>();
        var evidence = new List<object>();
        ValidateModel(document, selected, settings, diagnostics, evidence, cancellationToken);
        return new OfficeValidationReport(diagnostics, evidence);
    }

    private static OfficeDocumentCheck[] ValidationChecks(OfficeDocumentCheck[]? checks, OfficeDocumentValidationOptions options) {
        var selected = checks ?? new[] { OfficeDocumentCheck.Structure };
        if (selected.Length == 0) throw new ArgumentException("Select at least one validation check.", nameof(checks));
        foreach (var check in selected)
            if (!Enum.IsDefined(typeof(OfficeDocumentCheck), check)) throw new ArgumentOutOfRangeException(nameof(checks));
        if (options.Excel?.Doctor == null && options.Excel != null) throw new ArgumentException("Excel Doctor options are required.", nameof(options));
        if (options.Excel?.Doctor.RepairDefinedNames == true)
            throw new ArgumentException("Validation never repairs defined names. Use the repair operation instead.", nameof(options));
        if (options.Pdf?.DifferenceDocument != null && options.Pdf.DifferencePath != null)
            throw new ArgumentException("Specify DifferenceDocument or DifferencePath, not both.", nameof(options));
        return (options.Untrusted ? selected.Concat(new[] { OfficeDocumentCheck.Security }) : selected).Distinct().ToArray();
    }

    private static OfficePackageSecurityOptions ValidationSecurity(OfficeDocumentValidationOptions options) {
        var security = options.Security?.Snapshot() ?? OfficePackageSecurityOptions.SecureDefaults;
        if (options.Untrusted) {
            var untrusted = OfficePackageSecurityOptions.UntrustedDefaults;
            security.Macros = untrusted.Macros;
            security.ActiveX = untrusted.ActiveX;
            security.EmbeddedPayloads = untrusted.EmbeddedPayloads;
            security.ExternalRelationships = untrusted.ExternalRelationships;
        }
        return security;
    }

    private static void AddValidationSecurity(OfficePackageSecurityReport report, List<OfficeDiagnostic> diagnostics, List<object> evidence) {
        evidence.Add(report);
        foreach (var finding in report.Findings) diagnostics.Add(new OfficeDiagnostic("Security." + finding.Rule,
            (OfficeDiagnosticSeverity)(int)finding.Severity, finding.Message, finding.PartName, report.ContainerKind.ToString()));
    }

    private static void ValidateModel(object document, OfficeDocumentCheck[] checks, OfficeDocumentValidationOptions options,
        List<OfficeDiagnostic> diagnostics, List<object> evidence, CancellationToken cancellationToken, string? sourcePath = null,
        PdfLoadOptions? pdfRead = null) {
        PdfLoadOptions? effectivePdfRead = document is PdfDocument pdfModel ? pdfRead ?? ValidationPdfRead(options, pdfModel) : null;
        foreach (var check in checks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (document is PdfDocument pdf) { ValidatePdf(pdf, check, options, effectivePdfRead!, diagnostics, evidence, cancellationToken, sourcePath); continue; }
            switch (check) {
                case OfficeDocumentCheck.Structure:
                case OfficeDocumentCheck.Schema:
                    if (document is ExcelDocument excel && check == OfficeDocumentCheck.Structure)
                        ValidateExcel(excel, options.Excel, diagnostics, evidence, cancellationToken);
                    else {
                        IReadOnlyList<OfficeOpenXmlValidationError>? errors = document switch {
                            WordDocument word => word.ValidateDocument(),
                            ExcelDocument schemaWorkbook => schemaWorkbook.ValidateDocument(),
                            PowerPointDocument schemaPresentation => schemaPresentation.ValidateDocument(),
                            _ => null
                        };
                        if (errors == null) ValidationNotApplicable(diagnostics, check, document.GetType().Name);
                        else {
                            evidence.Add(errors);
                            foreach (var error in errors) diagnostics.Add(new OfficeDiagnostic("OpenXml." + error.Id,
                                OfficeDiagnosticSeverity.Error, error.Description, error.PartUri + error.Path, document.GetType().Name));
                        }
                    }
                    break;
                case OfficeDocumentCheck.Accessibility:
                    if (document is ExcelDocument workbook) {
                        var report = workbook.AnalyzeAccessibility();
                        evidence.Add(report);
                        foreach (var finding in report.Findings) diagnostics.Add(new OfficeDiagnostic("Excel.Accessibility." + finding.Category,
                            (OfficeDiagnosticSeverity)(int)finding.Severity, finding.Message, finding.SheetName + ":" + finding.Address, "Excel"));
                    } else if (document is PowerPointDocument presentation) {
                        var report = presentation.InspectAccessibility(options.PowerPointAccessibility);
                        evidence.Add(report);
                        foreach (var finding in report.Findings) diagnostics.Add(new OfficeDiagnostic(finding.Code,
                            (OfficeDiagnosticSeverity)(int)finding.Severity, finding.Message, "slide:" + finding.SlideIndex + "/shape:" + finding.ShapeId, "PowerPoint"));
                    } else ValidationNotApplicable(diagnostics, check, document.GetType().Name);
                    break;
                default:
                    ValidationNotApplicable(diagnostics, check, document.GetType().Name, check == OfficeDocumentCheck.Security
                        ? "Package security requires source bytes. A borrowed model is not saved or reopened for inspection." : null);
                    break;
            }
        }
    }

    private static void ValidationNotApplicable(List<OfficeDiagnostic> diagnostics, OfficeDocumentCheck check, string format, string? reason = null) =>
        diagnostics.Add(new OfficeDiagnostic("Validation.NotApplicable", OfficeDiagnosticSeverity.Information,
            check + " is not applicable to this input. " + reason, check.ToString(), format));
}
