namespace OfficeIMO.IWork;

/// <summary>Alignment of text at a recovered custom tab stop.</summary>
public enum IWorkTabAlignment {
    /// <summary>Aligns the left edge.</summary>
    Left,
    /// <summary>Centers text.</summary>
    Center,
    /// <summary>Aligns the right edge.</summary>
    Right,
    /// <summary>Aligns the decimal separator.</summary>
    Decimal
}

/// <summary>A qualified explicit tab stop without a leader.</summary>
public sealed class IWorkTabStop {
    internal IWorkTabStop(double positionPoints, IWorkTabAlignment alignment) {
        PositionPoints = positionPoints;
        Alignment = alignment;
    }

    /// <summary>Creates a finite nonnegative native tab stop without a leader, using native float precision.</summary>
    public static IWorkTabStop Create(float positionPoints, IWorkTabAlignment alignment = IWorkTabAlignment.Left) {
        IWorkCreationValidation.ValidateFinite(positionPoints, nameof(positionPoints));
        if (positionPoints < 0) throw new ArgumentOutOfRangeException(nameof(positionPoints));
        if ((uint)alignment > (uint)IWorkTabAlignment.Decimal) throw new ArgumentOutOfRangeException(nameof(alignment));
        return new IWorkTabStop(positionPoints == 0 ? 0 : positionPoints, alignment);
    }
    /// <summary>Gets the source position in points.</summary>
    public double PositionPoints { get; }
    /// <summary>Gets the text alignment.</summary>
    public IWorkTabAlignment Alignment { get; }
}
