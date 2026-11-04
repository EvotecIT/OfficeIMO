using System.Net.Http;

namespace OfficeIMO.Excel;

public sealed partial class ExcelCreateOptions {
    /// <summary>Optional visual theme for a new workbook. Null retains the workbook or template defaults.</summary>
    public OfficeTheme? Theme { get; set; }
    /// <summary>Optional workbook template copied into the new document.</summary>
    public string? Template { get; set; }
    /// <summary>Optional caller-owned template stream. It cannot be combined with Template.</summary>
    public Stream? TemplateStream { get; set; }
    /// <summary>Allows a template destination to replace an existing file. The default is true.</summary>
    public bool Overwrite { get; set; } = true;
    /// <summary>Optional package settings applied when reading the template.</summary>
    public OfficeOpenXmlLoadSettings? OpenSettings { get; set; }
    /// <summary>Optional worksheet created when a new workbook has no sheets.</summary>
    public string? FirstSheetName { get; set; }
}

public sealed partial class ExcelLoadOptions {
    /// <summary>Password for an encrypted workbook source.</summary>
    public string? Password { get; set; }
    /// <summary>Stages an Open XML workbook in a bounded temporary file instead of memory.</summary>
    public bool FileBacked { get; set; }
    /// <summary>Maximum binary XLS input bytes admitted by the native decoder.</summary>
    public int MaxLegacyInputBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum aggregate decoded legacy picture bytes.</summary>
    public int MaxDecodedImageBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Reports recognized unsupported legacy workbook features.</summary>
    public bool ReportUnsupportedContent { get; set; } = true;
    /// <summary>Retains active external-workbook links in legacy projections. The default keeps them inert.</summary>
    public bool PreserveExternalWorkbookLinks { get; set; }

    internal LegacyXls.LegacyXlsImportOptions CreateLegacyXlsOptions() => new LegacyXls.LegacyXlsImportOptions {
        MaxInputBytes = (int)Math.Min(MaxInputBytes ?? int.MaxValue, MaxLegacyInputBytes),
        MaxDecodedImageBytes = MaxDecodedImageBytes,
        ReportUnsupportedContent = ReportUnsupportedContent,
        PreserveExternalWorkbookLinks = PreserveExternalWorkbookLinks,
        Password = Password
    };
}
