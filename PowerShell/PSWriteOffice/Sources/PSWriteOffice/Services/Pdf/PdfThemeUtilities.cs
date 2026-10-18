using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Cmdlets.Pdf;

namespace PSWriteOffice.Services.Pdf;

internal static class PdfThemeUtilities
{
    internal static PdfTheme ResolveTheme(OfficePdfThemePreset theme)
    {
        return theme switch
        {
            OfficePdfThemePreset.WordLike => PdfTheme.WordLike(),
            OfficePdfThemePreset.TechnicalDocument => PdfTheme.TechnicalDocument(),
            OfficePdfThemePreset.Compact => PdfTheme.Compact(),
            OfficePdfThemePreset.Report => PdfTheme.Report(),
            _ => throw new PSArgumentOutOfRangeException(nameof(theme), theme, "Unsupported PDF theme preset.")
        };
    }
}
