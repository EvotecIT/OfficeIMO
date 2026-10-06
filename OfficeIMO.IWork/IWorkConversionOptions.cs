namespace OfficeIMO.IWork;

/// <summary>Controls how an opened iWork source is represented by a destination adapter.</summary>
public sealed class IWorkConversionOptions {
    /// <summary>Gets or sets whether conversion prefers editable content or requires a particular representation.</summary>
    public IWorkConversionMode Mode { get; set; } = IWorkConversionMode.Auto;

    /// <summary>Gets or sets whether an adapter may keep recovered editable objects while reporting incomplete source or destination details.</summary>
    /// <remarks>Defaults to false. Package limits and destination bounds still apply. This does not promise complete content or visual fidelity.</remarks>
    public bool AllowPartialEditableReconstruction { get; set; }

    /// <summary>Gets or sets whether visual fallback must use an asset known to cover the complete source.</summary>
    /// <remarks>Defaults to true. Set this to false explicitly when a first-page or composite preview is acceptable, and retain the conversion result to inspect its coverage report.</remarks>
    public bool RequireCompleteVisualCoverage { get; set; } = true;

    /// <summary>Gets or sets whether the Numbers adapter may normalize destination worksheet names using the Excel owner's collision-safe rules.</summary>
    /// <remarks>Defaults to false. Renames are reported and the original source sheet/table identity remains available on the conversion result.</remarks>
    public bool NormalizeWorksheetNames { get; set; }

    /// <summary>Returns a known destination omission that requires the explicit partial policy.</summary>
    internal string? FindUnacceptedDestinationOmission(IEnumerable<IWorkDiagnostic> diagnostics) =>
        AllowPartialEditableReconstruction ? null : diagnostics.FirstOrDefault(diagnostic =>
            diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission)?.Message;

    internal void ValidateVisualPreview(IWorkPreviewAsset? preview) {
        if (RequireCompleteVisualCoverage && preview?.Coverage != IWorkVisualCoverage.FullDocument) {
            throw new InvalidDataException("The source preview is not known to cover the complete document. Use editable reconstruction or explicitly permit incomplete visual coverage.");
        }
    }

    /// <summary>Validates these options and returns an independent copy for one conversion.</summary>
    public IWorkConversionOptions Clone() {
        if (Mode is not (IWorkConversionMode.Auto
                or IWorkConversionMode.EditableOnly
                or IWorkConversionMode.VisualOnly)) {
            throw new ArgumentOutOfRangeException(nameof(Mode),
                "The conversion mode is not defined.");
        }
        return (IWorkConversionOptions)MemberwiseClone();
    }
}
