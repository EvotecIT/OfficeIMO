namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>High-level PowerPoint shape categories returned by the PSWriteOffice shape reader.</summary>
public enum OfficePowerPointShapeKind
{
    /// <summary>Text box or placeholder.</summary>
    TextBox,
    /// <summary>Picture shape.</summary>
    Picture,
    /// <summary>Table shape.</summary>
    Table,
    /// <summary>Chart shape.</summary>
    Chart,
    /// <summary>Preset auto shape.</summary>
    AutoShape,
    /// <summary>Grouped shape.</summary>
    GroupShape
}
