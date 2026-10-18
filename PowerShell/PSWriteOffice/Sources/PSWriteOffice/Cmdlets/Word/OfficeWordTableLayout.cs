namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Word table layout and auto-fit behaviors exposed by <c>Add-OfficeWordTable</c>.</summary>
public enum OfficeWordTableLayout
{
    /// <summary>Allow Word to adjust columns based on content.</summary>
    AutoFit,
    /// <summary>Use fixed preferred widths.</summary>
    Fixed,
    /// <summary>Run Word's auto-fit-to-contents operation after table creation.</summary>
    AutoFitToContents,
    /// <summary>Run Word's auto-fit-to-window operation after table creation.</summary>
    AutoFitToWindow
}
