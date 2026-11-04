namespace PSWriteOffice.Models.Excel;

/// <summary>A manual print page-break snapshot that remains readable after its source workbook is closed.</summary>
public sealed class ExcelPageBreakInfo {
    internal ExcelPageBreakInfo(string type, int position, string sheetName, string? path) {
        Type = type; Position = position; SheetName = sheetName; Path = path;
    }
    /// <summary>Row or Column.</summary>
    public string Type { get; }
    /// <summary>Alias of Type.</summary>
    public string Kind => Type;
    /// <summary>One-based row or column after which printing starts a new page.</summary>
    public int Position { get; }
    /// <summary>Alias of Position.</summary>
    public int Index => Position;
    /// <summary>Worksheet containing the break.</summary>
    public string SheetName { get; }
    /// <summary>Alias of SheetName.</summary>
    public string Sheet => SheetName;
    /// <summary>Source file, or null for a caller-owned workbook.</summary>
    public string? Path { get; }
    /// <summary>Alias of Path.</summary>
    public string? InputPath => Path;
}
