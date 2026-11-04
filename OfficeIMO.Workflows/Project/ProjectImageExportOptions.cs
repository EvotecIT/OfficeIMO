using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

/// <summary>Project report image settings. Pages use points; the default PNG density is 300 DPI.</summary>
public sealed class ProjectImageExportOptions : OfficeImageExportOptions {
    /// <summary>Creates detailed report defaults. Explicit quality, DPI and scale settings remain available.</summary>
    public ProjectImageExportOptions() {
        TargetDpi = 300;
        RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw;
        TextShapingProvider = OfficeManagedTextShapingProvider.Instance;
    }
    /// <summary>Project drawings use 72 points per inch.</summary>
    public override double LogicalUnitsPerInch => 72;
    internal ProjectImageExportOptions Snapshot() => CopyImageExportOptionsTo(new ProjectImageExportOptions());
}
