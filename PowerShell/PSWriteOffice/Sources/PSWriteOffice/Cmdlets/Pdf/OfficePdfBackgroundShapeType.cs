namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Supported generated PDF background shape types.</summary>
public enum OfficePdfBackgroundShapeType
{
    /// <summary>Rectangle at explicit page coordinates.</summary>
    Rectangle,

    /// <summary>Rounded rectangle at explicit page coordinates.</summary>
    RoundedRectangle,

    /// <summary>Ellipse at explicit page coordinates.</summary>
    Ellipse,

    /// <summary>Full-width band anchored to the top of the page.</summary>
    TopBand,

    /// <summary>Full-width band anchored to the bottom of the page.</summary>
    BottomBand,

    /// <summary>Full-height band anchored to the left of the page.</summary>
    LeftBand,

    /// <summary>Full-height band anchored to the right of the page.</summary>
    RightBand
}
