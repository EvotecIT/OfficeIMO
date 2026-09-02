namespace OfficeIMO.IWork;

/// <summary>Controls how an opened iWork source is represented by a destination adapter.</summary>
public sealed class IWorkConversionOptions {
    /// <summary>Gets or sets whether conversion prefers editable content or requires a particular representation.</summary>
    public IWorkConversionMode Mode { get; set; } = IWorkConversionMode.Auto;

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
