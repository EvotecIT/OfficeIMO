using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO;

public static partial class OfficeDocument {
    private static readonly string[] AccessibilityExtensions = {
        ".docx", ".docm", ".dotx", ".dotm", ".doc", ".dot",
        ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlt", ".xlsb",
        ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".ppt", ".pps", ".pot"
    };

    /// <summary>Returns whether <see cref="InspectAccessibility(object, OfficeAccessibilityPolicyProfile, string?, CancellationToken)"/> supports the file extension.</summary>
    public static bool SupportsAccessibility(string path) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return Array.IndexOf(AccessibilityExtensions, Path.GetExtension(path).ToLowerInvariant()) >= 0;
    }

    /// <summary>
    /// Inspects a Word, Excel or PowerPoint file or open model against the built-in accessibility policy.
    /// Every format returns the same report, finding codes and severities, so one CI gate covers all three.
    /// </summary>
    /// <param name="document">A file path, or an open <see cref="WordDocument"/>, <see cref="ExcelDocument"/> or <see cref="PowerPointDocument"/>.
    /// Paths open read-only; legacy binary files are inspected through their loaded model. Open models are not changed.</param>
    /// <param name="profile">Built-in policy profile. Strict turns structural warnings into errors and requires a document title.</param>
    /// <param name="password">Password used only when opening an encrypted path input.</param>
    /// <param name="cancellationToken">Cancels loading a path input.</param>
    public static OfficeAccessibilityReport InspectAccessibility(object document,
        OfficeAccessibilityPolicyProfile profile = OfficeAccessibilityPolicyProfile.Default,
        string? password = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        if (document is string path) {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);
            if (!File.Exists(path)) throw new FileNotFoundException("Document not found.", path);
            string extension = Path.GetExtension(path).ToLowerInvariant();
            using IDisposable model = LoadOfficeModelReadOnly(path, extension, password, cancellationToken)
                ?? throw new NotSupportedException("Accessibility inspection supports Word, Excel and PowerPoint files, not '" + extension + "'.");
            cancellationToken.ThrowIfCancellationRequested();
            return InspectAccessibility(model, profile);
        }
        return document switch {
            WordDocument word => word.InspectAccessibility(new WordAccessibilityOptions { Profile = profile }),
            ExcelDocument workbook => workbook.InspectAccessibility(ExcelAccessibilityOptions.ForProfile(profile)),
            PowerPointDocument presentation => presentation.InspectAccessibility(PowerPointAccessibilityOptions.ForProfile(profile)),
            _ => throw new NotSupportedException("Accessibility inspection supports WordDocument, ExcelDocument and PowerPointDocument, not " + document.GetType().Name + ".")
        };
    }
}
