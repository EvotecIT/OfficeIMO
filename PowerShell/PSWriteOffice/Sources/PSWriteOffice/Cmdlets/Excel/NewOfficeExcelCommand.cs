using System;
using System.Collections;
using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates a themed Excel workbook, returning a live workbook or saving a completed content block.</summary>
/// <para>Without a block no file is written, including on disposal. A block saves atomically; an existing destination requires Force. Template copies content; Theme reads only the visual theme. SaveOptions and Password apply to a completed content block.</para>
/// <example><summary>Create a workbook and return its saved file.</summary>
/// <code>New-OfficeExcel -Path ./Report.xlsx -Theme Modern -PassThru {
///     ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Region' }
/// }</code></example>
/// <example><summary>Create a live workbook for explicit editing and saving.</summary>
/// <code>$workbook = New-OfficeExcel -Theme Boardroom
/// $sheet = Add-OfficeExcelSheet -Document $workbook -Name Data -PassThru
/// Set-OfficeExcelCell -Sheet $sheet -Address A1 -Value 'Region'
/// Save-OfficeExcel -Document $workbook -Path ./Report.xlsx
/// Close-OfficeExcel -Document $workbook</code></example>
[Cmdlet(VerbsCommon.New, "OfficeExcel", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(ExcelDocument), typeof(FileInfo))]
public sealed class NewOfficeExcelCommand : OfficeNewLifecycleCmdlet<ExcelDocument> {
    /// <summary>Built-in theme, theme object or a document/template whose visual theme is read.</summary>
    [Parameter, OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Existing workbook copied as starting content. Its styling is preserved unless Theme is supplied.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Template { get; set; }

    /// <summary>Allow a completed content block to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Password for a completed content block's saved file; use Save-OfficeExcel for a live workbook.</summary>
    [Parameter(ParameterSetName = "Content"), OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Native save options or a hashtable, applied to a completed content block.</summary>
    [Parameter(ParameterSetName = "Content"), OfficeExcelSaveOptionsTransformation, ValidateNotNull]
    public ExcelSaveOptions? SaveOptions { get; set; }

    /// <summary>Workbook properties such as Title, Author, Subject, Keywords, Company and Manager.</summary>
    [Parameter, ValidateNotNull]
    public Hashtable? DocumentProperty { get; set; }

    /// <summary>Native workbook date system.</summary>
    [Parameter]
    public ExcelDateSystem? DateSystem { get; set; }

    /// <summary>Maximum parallel workers; 1 runs sequentially. Omission retains the native execution policy.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxDegreeOfParallelism { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ExcelCreateOptions {
            PersistenceMode = DocumentPersistenceMode.Explicit,
            Theme = Theme?.Resolve(OfficeTheme.FromDocument),
            Template = Template == null ? null : ResolveOutputPath(Template)
        };
        RunDocument(Path, isNew: true, force: Force.IsPresent, path => {
            ExcelDocument document = path == null ? ExcelDocument.Create(options) : ExcelDocument.Create(path, options);
            try {
                if (DateSystem.HasValue) document.DateSystem = DateSystem.Value;
                if (MaxDegreeOfParallelism.HasValue) document.Execution.MaxDegreeOfParallelism = MaxDegreeOfParallelism.Value;
                if (DocumentProperty != null) {
                    foreach (DictionaryEntry property in DocumentProperty)
                        ExcelDocumentPropertyService.SetProperty(document, Convert.ToString(property.Key)!, property.Value);
                }
                return document;
            } catch { document.Dispose(); throw; }
        }, (document, path, allowLoss) => {
            if (document.Sheets.Count == 0) document.AddSheet("Sheet1");
            document.Save(path, ExcelDocumentService.NativeSaveOptions(SaveOptions,
                MyInvocation.BoundParameters.ContainsKey(nameof(AllowLoss)) ? allowLoss : (bool?)null,
                Password, MyInvocation.BoundParameters.ContainsKey(nameof(Password))));
            return null;
        });
    }

    /// <inheritdoc />
    protected override void InvokeContent(ExcelDocument document) {
        using (ExcelDslContext.Enter(document)) Content!.InvokeReturnAsIs();
    }
}
