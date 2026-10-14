using System;

namespace OfficeIMO.Excel;

/// <summary>Independent rectangular cell edges selected for border formatting.</summary>
[Flags]
public enum ExcelCellBorders {
    /// <summary>No selected edges.</summary>
    None = 0,
    /// <summary>The left edge.</summary>
    Left = 1,
    /// <summary>The top edge.</summary>
    Top = 2,
    /// <summary>The right edge.</summary>
    Right = 4,
    /// <summary>The bottom edge.</summary>
    Bottom = 8,
    /// <summary>All four rectangular edges.</summary>
    All = Left | Top | Right | Bottom
}
