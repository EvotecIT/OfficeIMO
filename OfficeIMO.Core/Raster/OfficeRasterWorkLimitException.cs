namespace OfficeIMO.Drawing;

/// <summary>Raised when one element's coverage work exceeds the rasterizer's bounded per-element budget.</summary>
internal sealed class OfficeRasterWorkLimitException : InvalidOperationException {
    /// <summary>Diagnostic code recorded when a render omits an element for this reason.</summary>
    internal const string DiagnosticCode = "RasterElementWorkLimit";

    internal OfficeRasterWorkLimitException(string message) : base(message) { }
}
