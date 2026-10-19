using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Opens an existing Excel workbook.</summary>
/// <para>Returns the underlying <see cref="ExcelDocument"/> so callers can inspect or reuse it in DSL pipelines.</para>
/// <example>
///   <summary>Load a workbook in read-only mode.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook = Get-OfficeExcel -Path .\report.xlsx -ReadOnly</code>
///   <para>Loads <c>report.xlsx</c> for inspection without enabling writes.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcel", DefaultParameterSetName = ParameterSetPath)]
public sealed class GetOfficeExcelCommand : AsyncPSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetUri = "Uri";

    /// <summary>Path to the workbook to load.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Remote workbook URI to load.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetUri)]
    [Alias("Url")]
    public Uri? Uri { get; set; }

    /// <summary>Allow HTTP workbook downloads in addition to HTTPS.</summary>
    [Parameter(ParameterSetName = ParameterSetUri)]
    public SwitchParameter AllowHttp { get; set; }

    /// <summary>Open the file in read-only mode.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Password used to open an encrypted workbook package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ParameterSetName == ParameterSetUri) {
            if (Uri == null) {
                throw new PSArgumentException("Workbook URI was not provided.", nameof(Uri));
            }

            var remoteDocument = await ExcelDocumentService.LoadDocumentAsync(
                Uri,
                ReadOnly.IsPresent,
                AllowHttp.IsPresent,
                Password,
                CancelToken).ConfigureAwait(false);
            WriteObject(remoteDocument);
            return;
        }

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!File.Exists(resolvedPath)) {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        ExcelDocument document = ExcelDocumentService.LoadDocument(resolvedPath, ReadOnly.IsPresent, autoSave: false, Password);
        WriteObject(document);
    }
}
