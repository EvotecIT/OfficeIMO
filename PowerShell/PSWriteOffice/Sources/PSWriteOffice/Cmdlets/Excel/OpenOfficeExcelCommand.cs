using System;
using System.IO;
using System.Management.Automation;
using System.Net;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Opens a workbook for explicit editing, or atomically saves a completed content block.</summary>
/// <para>Opening preserves the workbook's theme and encryption. A failed block leaves its source unchanged. Remote workbooks remain detached; save them with Save-OfficeExcel -Path.</para>
/// <example><summary>Open a literal file from the pipeline.</summary>
/// <code>$workbook = Get-Item -LiteralPath './Review[final].xlsx' | Open-OfficeExcel
/// $workbook | Close-OfficeExcel</code></example>
/// <example><summary>Edit and save an existing workbook.</summary>
/// <code>Open-OfficeExcel -Path ./Report.xlsx {
///     ExcelSheet 'Data' { ExcelCell -Address B2 -Value 'Reviewed' }
/// }</code></example>
/// <example><summary>Read a detached workbook from HTTPS.</summary>
/// <code>$workbook = Open-OfficeExcel -Uri https://intranet.example.com/Budget.xlsx -ReadOnly
/// Import-OfficeExcel -Document $workbook -Sheet Budget
/// Close-OfficeExcel -Document $workbook</code></example>
[Cmdlet(VerbsCommon.Open, "OfficeExcel", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(ExcelDocument), typeof(FileInfo))]
public sealed class OpenOfficeExcelCommand : AsyncPSCmdlet {
    /// <summary>Workbook path with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Literal workbook path; accepts FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;

    /// <summary>HTTP or HTTPS workbook loaded without a local save destination.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Uri"), ValidateNotNull]
    public Uri Uri { get; set; } = null!;

    /// <summary>Allow HTTP in addition to HTTPS for a remote workbook.</summary>
    [Parameter(ParameterSetName = "Uri")]
    public SwitchParameter AllowHttp { get; set; }

    /// <summary>Edit a local workbook and save only after the block completes successfully.</summary>
    [Parameter(Position = 1, ParameterSetName = "Path"), Parameter(Position = 1, ParameterSetName = "LiteralPath")]
    [ValidateNotNull]
    public ScriptBlock? Content { get; set; }

    /// <summary>Return a read-only live workbook; authoring blocks require writable input.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Source password as a SecureString, credential, or warned plaintext value.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Native save options or a hashtable, applied only to a completed local content block.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    [OfficeExcelSaveOptionsTransformation, ValidateNotNull]
    public ExcelSaveOptions? SaveOptions { get; set; }

    /// <summary>Explicitly accept native-save reductions in a completed local content block.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Maximum parallel workers; omission preserves the native execution policy.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>Return the saved FileInfo from a completed local content block.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Open the saved file in its registered application after a local block completes.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Show { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ReadOnly && Content != null)
            throw new PSArgumentException("ReadOnly cannot be combined with an authoring content block.");
        if (Content == null && (SaveOptions != null || MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss)) || PassThru || Show))
            throw new PSArgumentException("SaveOptions, AllowLoss, PassThru and Show require a local content block; use Save-OfficeExcel for a live workbook.");

        string? password = Password == null ? null : new NetworkCredential(string.Empty, Password).Password;
        if (ParameterSetName == "Uri") {
            ExcelDocument remote = await ExcelDocumentService.LoadDocumentAsync(Uri, ReadOnly.IsPresent,
                AllowHttp.IsPresent, password, CancelToken).ConfigureAwait(false);
            ReturnLive(remote);
            return;
        }

        // The local branch does not await: authoring stays on the PowerShell pipeline thread.
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInput(SessionState, literal ? LiteralPath : Path, literal)) {
            if (Content == null) {
                ReturnLive(ExcelDocumentService.LoadDocument(source, ReadOnly.IsPresent, autoSave: false, password));
                continue;
            }
            if (!ShouldProcess(source, "Save Excel workbook")) continue;
            ExcelSaveOptions options = ExcelDocumentService.NativeSaveOptions(SaveOptions,
                MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss)) ? AllowLoss.IsPresent : (bool?)null,
                password: null, passwordBound: false);
            FileInfo result = OfficeContentLifecycle.Run(source, isNew: false, refuseExisting: false, options.AllowLoss,
                path => ApplyExecution(ExcelDocumentService.LoadDocument(path, readOnly: false, autoSave: false, password)),
                document => { using (ExcelDslContext.Enter(document)) Content!.InvokeReturnAsIs(); },
                (document, path, _) => { document.Save(path, options); return null; }, CancelToken,
                error => ThrowTerminatingError(OfficeError.NativeSave(error, source)));
            if (Show) OfficeShow.Open(result.FullName);
            if (PassThru) WriteObject(result);
        }
    }

    private ExcelDocument ApplyExecution(ExcelDocument document) {
        try {
            if (MaxDegreeOfParallelism.HasValue) document.Execution.MaxDegreeOfParallelism = MaxDegreeOfParallelism.Value;
            return document;
        } catch { document.Dispose(); throw; }
    }

    private void ReturnLive(ExcelDocument document) {
        ApplyExecution(document);
        try { WriteObject(document); }
        catch { document.Dispose(); throw; }
    }
}
