using OfficeIMO.Excel;

namespace PSWriteOffice.Models;

/// <summary>Native workbook repair findings and the destination actually published by Repair-OfficeDocument.</summary>
public sealed class OfficeRepairReport {
    internal OfficeRepairReport(ExcelWorkbookRepairReport nativeReport, string? inputPath, string? outputPath) {
        NativeReport = nativeReport; InputPath = inputPath; OutputPath = outputPath;
    }
    /// <summary>Complete native actions and before/after diagnostics.</summary>
    public ExcelWorkbookRepairReport NativeReport { get; }
    /// <summary>Source path, or null for a caller-owned model.</summary>
    public string? InputPath { get; }
    /// <summary>Published destination, or null when only a live model changed.</summary>
    public string? OutputPath { get; }
    /// <summary>Number of actions the native repair owner attempted.</summary>
    public int ActionCount => NativeReport.ActionCount;
}
