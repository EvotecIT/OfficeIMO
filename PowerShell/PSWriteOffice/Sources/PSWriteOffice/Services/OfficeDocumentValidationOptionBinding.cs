using System.Collections;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

/// <summary>Routes complete native check settings through the shared options binder; later values replace their matching settings group.</summary>
internal static class OfficeDocumentValidationOptionBinding {
    internal static OfficeDocumentValidationOptions Capture(EngineIntrinsics engine, object input, object[]? values, OfficeDocumentCheck[] checks) {
        string extension = input is string path ? System.IO.Path.GetExtension(path).ToLowerInvariant() : string.Empty;
        bool excel = input is ExcelDocument || new[] { ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlt", ".xlsb" }.Contains(extension);
        bool pdf = input is PdfDocument || extension == ".pdf";
        bool word = input is WordDocument || new[] { ".docx", ".docm", ".dotx", ".dotm", ".doc", ".dot" }.Contains(extension);
        bool presentation = input is PowerPointDocument || new[] { ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".ppt", ".pot", ".pps" }.Contains(extension);
        var result = new OfficeDocumentValidationOptions();
        foreach (object item in values ?? Array.Empty<object>()) {
            object? value = OfficeArgumentValue.Unwrap(item);
            if (value is Models.Pdf.PdfReadOptions shellPdf && pdf) {
                var captured = shellPdf.Snapshot();
                if (captured.OcrMergeOptions != null)
                    throw new PSArgumentException("OCR options apply to Read-OfficeDocument -As Ocr, not document validation.");
                value = captured.LoadOptions;
            }
            if (value is ExcelTestOptions shell && excel) value = new OfficeExcelValidationOptions {
                Doctor = shell.Doctor, Capability = shell.Capability, IncludeFeatures = shell.IncludeFeatures, IncludeRepairHints = shell.IncludeRepairHints
            };
            // A fresh holder ensures a hashtable never modifies an earlier caller-owned typed options object.
            var holder = new Slots();
            var names = new HashSet<string> { nameof(Slots.Security) };
            if (excel) { names.Add(nameof(Slots.Excel)); names.Add(nameof(Slots.Doctor)); names.Add(nameof(Slots.ExcelAccessibility)); }
            if (word) names.Add(nameof(Slots.WordAccessibility));
            if (pdf) { names.Add(nameof(Slots.Pdf)); names.Add(nameof(Slots.PdfRead)); names.Add(nameof(Slots.Rewrite)); }
            if (presentation) names.Add(nameof(Slots.PowerPointAccessibility));
            if (pdf && value is IDictionary table && table.Keys.Cast<object>().Any(key =>
                key is string name && name.Equals("Profile", StringComparison.OrdinalIgnoreCase))) {
                // Reading and compliance use unrelated native Profile enums. The selected check owns a flat Profile;
                // a typed PdfLoadOptions remains available when both policies are needed in one invocation.
                bool compliance = checks.Contains(OfficeDocumentCheck.Compliance) || checks.Contains(OfficeDocumentCheck.Accessibility);
                names.Remove(compliance ? nameof(Slots.PdfRead) : nameof(Slots.Pdf));
            }
            var slots = typeof(Slots).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property => names.Contains(property.Name)).Select(property => ((object)holder, property)).ToArray();
            OfficeOptionPropertyBinding.Apply(engine, value, slots);
            if (holder.Security != null) result.Security = holder.Security;
            if (holder.Excel != null) result.Excel = CopyExcel(holder.Excel);
            if (holder.Doctor != null) {
                result.Excel = result.Excel == null ? new OfficeExcelValidationOptions() : CopyExcel(result.Excel);
                result.Excel.Doctor = holder.Doctor;
            }
            if (holder.PdfRead != null) result.PdfRead = holder.PdfRead.Clone();
            if (holder.Pdf != null) result.Pdf = CopyPdf(holder.Pdf);
            if (holder.Rewrite != null) {
                result.Pdf = result.Pdf == null ? new OfficePdfValidationOptions() : CopyPdf(result.Pdf);
                result.Pdf.Rewrite = holder.Rewrite;
            }
            if (holder.PowerPointAccessibility != null) result.PowerPointAccessibility = holder.PowerPointAccessibility;
            if (holder.WordAccessibility != null) result.WordAccessibility = holder.WordAccessibility;
            if (holder.ExcelAccessibility != null) result.ExcelAccessibility = holder.ExcelAccessibility;
        }
        return result;
    }

    private static OfficeExcelValidationOptions CopyExcel(OfficeExcelValidationOptions source) => new() {
        Doctor = source.Doctor, Capability = source.Capability, IncludeFeatures = source.IncludeFeatures, IncludeRepairHints = source.IncludeRepairHints
    };
    private static OfficePdfValidationOptions CopyPdf(OfficePdfValidationOptions source) => new() {
        Profile = source.Profile, Proof = source.Proof, ExternalValidation = source.ExternalValidation,
        Capability = source.Capability, DifferenceDocument = source.DifferenceDocument, DifferencePath = source.DifferencePath, Rewrite = source.Rewrite
    };

    private sealed class Slots {
        public OfficePackageSecurityOptions? Security { get; set; }
        public OfficeExcelValidationOptions? Excel { get; set; }
        public ExcelWorkbookDoctorOptions? Doctor { get; set; }
        public PdfLoadOptions? PdfRead { get; set; }
        public OfficePdfValidationOptions? Pdf { get; set; }
        public PdfRewritePreservationOptions? Rewrite { get; set; }
        public PowerPointAccessibilityOptions? PowerPointAccessibility { get; set; }
        public WordAccessibilityOptions? WordAccessibility { get; set; }
        public ExcelAccessibilityOptions? ExcelAccessibility { get; set; }
    }
}
