using System.Globalization;
using OfficeIMO.Internal;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class InvoiceWorkbenchViewModel {
    private OfficeInvoiceStorageWorkflowRequest CaptureRequest() {
        var operation = SelectedOperation.Value;
        InvoiceXmlOptions? target = NeedsTarget ? new(SelectedTarget.Options.Release, SelectedTarget.Options.Syntax, SelectedTarget.Options.Profile,
            AllowProfileLoss ? InvoiceProjectionPolicy.AllowProfileDefinedDataLoss : InvoiceProjectionPolicy.RejectDataLoss) : null;
        InvoiceSourceEdits? edits = null;
        if (IsEditing) {
            edits = new(Nonempty(EditNumber), Date(EditIssueDate), Date(EditDueDate), Nonempty(EditBuyerReference), Nonempty(EditPaymentReference));
            if (edits.Number == null && edits.IssueDate == null && edits.DueDate == null && edits.BuyerReference == null && edits.PaymentReference == null)
                throw new ArgumentException(T("EditRequired", "Enter at least one source-field replacement."));
        }
        var layout = IsRendering ? InvoicePdfLayoutOptions.ForCultures(Languages.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)) : new();
        PdfOptions? pdfOptions = null;
        if (IsRendering) {
            layout.LineColumns.Clear();
            foreach (var column in SelectedTableLayout.Columns) layout.LineColumns.Add(column);
            layout.UnitCodeDisplay = SelectedUnitDisplay.Value; layout.PaymentCodeDisplay = SelectedPaymentDisplay.Value;
            layout.CompactDetails = CompactDetails; layout.IncludePageIdentity = PageIdentity;
            if (ModernLayout) layout.Theme = new InvoicePdfTheme();
            if (!string.IsNullOrWhiteSpace(FontPath)) {
                using var file = File.OpenRead(FontPath);
                if (file.Length is <= 0 or > 16 * 1024 * 1024) throw new ArgumentException(T("FontLimit", "The embedded font must contain between one byte and 16 MiB."));
                byte[] font = new byte[(int)file.Length]; file.ReadExactly(font);
                pdfOptions = new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, font, "Invoice font")
                    .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Invoice font");
            }
        }
        string? output = null;
        if (IsWriting) {
            string folder = string.IsNullOrWhiteSpace(OutputFolder)
                ? Path.GetDirectoryName(OfficeStorageIdentity.GetLocalPath(InputPath) ?? throw new ArgumentException(T("OutputRequired", "Choose an output folder.")))!
                : OutputFolder;
            output = _storage?.UsesProviderPublication(folder) == true ? folder
                : Path.Combine(OfficeStorageIdentity.GetLocalPath(folder) ?? throw new ArgumentException(T("OutputRequired", "Choose an output folder.")), GetOutputName());
        }
        return new() {
            InputPath = InputPath, InputStream = _storage?.CreateWorkflowInput(InputPath), Operation = operation,
            OutputPath = output, Target = target, SourceEdits = edits,
            ValidationRelease = RequireStandards ? target?.Release ?? SelectedRelease.Value : null,
            Layout = layout.Clone(), PdfOptions = pdfOptions, PublicationGuard = _guard, ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
        };
    }
    private InvoiceValidator? CaptureValidator() {
        if (!RequireStandards) return null;
        if (string.IsNullOrWhiteSpace(RuleBundlePath) || string.IsNullOrWhiteSpace(SaxonJarPath))
            throw new ArgumentException(T("RulesRequired", "Select the pinned rule bundle and Saxon JAR before requesting standards checks."));
        return new(InvoiceRuleBundle.Load(RuleBundlePath, Nonempty(PeppolRulesPath), Nonempty(FacturXRulesPath)),
            new SaxonInvoiceRulesRunner(SaxonJarPath, Nonempty(JavaExecutable) ?? "java"));
    }
    private string GetOutputName() {
        string stem = Path.GetFileNameWithoutExtension(InputFileName);
        string suffix = SelectedOperation.Value switch {
            OfficeInvoiceWorkflowOperation.EditSource => ".edited.xml",
            OfficeInvoiceWorkflowOperation.Convert => ".converted.xml",
            OfficeInvoiceWorkflowOperation.RenderPresentationPdf => ".presentation.pdf",
            _ => ".invoice.pdf"
        };
        return stem + suffix;
    }
    private static string? Nonempty(string value) => string.IsNullOrWhiteSpace(value) ? null : value;
    private DateTime? Date(string value) => string.IsNullOrWhiteSpace(value) ? null :
        DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date) ? date
        : throw new ArgumentException(T("DateRequired", "Invoice dates require yyyy-MM-dd."));
    private void ShowReport(OfficeInvoiceStorageWorkflowResult result, bool standardsRequested) {
        Status = result.Summary; OutputPath = result.OutputPath;
        var report = result.Workflow;
        if (report?.Source is { } source) {
            string Short(string? text) => text == null ? "—" : text.Length <= 160 ? text : text.Substring(0, 160) + "…";
            SourceSummary = string.Join(" · ", Short(source.Invoice.Number), Short(source.Invoice.Seller?.Name), Short(source.Invoice.Buyer?.Name), Short(source.Invoice.Currency));
            MappingSummary = source.HasCompleteMapping ? T("Mapping.Complete", "Complete semantic mapping") : T("Mapping.Partial", "Source contains unmapped data; inspect the findings.");
        }
        ModelSummary = report?.ModelValidation is { } model ? model.IsValid ? T("Model.Valid", "Model checks passed")
            : T("Model.Invalid", "Model checks found invalid business data") : T("Model.Unavailable", "Model checks unavailable");
        SchemaSummary = FormatStandardsStatus(report?.SchemaStatus ?? InvoiceValidationStatus.NotRun, standardsRequested);
        RulesSummary = FormatStandardsStatus(report?.BusinessRulesStatus ?? InvoiceValidationStatus.NotRun, standardsRequested);
        foreach (var finding in report?.Diagnostics ?? []) Diagnostics.Add($"{finding.Severity} · {finding.Code} · {finding.Location}: {finding.Message}");
        foreach (var finding in result.Diagnostics) Diagnostics.Add($"{finding.Severity} · {finding.Code}: {finding.Message}");
        if (result.Recovery != null) Diagnostics.Add(T("Recovery", "A verified recovery copy is available in Jobs. Check it and the provider destination before retrying."));
    }

    internal string FormatStandardsStatus(InvoiceValidationStatus status, bool requested) =>
        status == InvoiceValidationStatus.NotRun && !requested ? T("ValidationStatus.NotRequested", "Not requested")
            : T("ValidationStatus." + status, status == InvoiceValidationStatus.NotRun ? "Not run" : status.ToString());
}
