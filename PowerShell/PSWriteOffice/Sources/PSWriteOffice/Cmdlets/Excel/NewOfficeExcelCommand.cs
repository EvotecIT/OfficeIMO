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
/// <example>
/// <summary>Create a workbook with one table</summary>
/// <code>New-OfficeExcel -Path .\Services.xlsx {
///     ExcelSheet 'Services' {
///         ExcelTable (Get-Service | Select-Object Name, DisplayName, Status, StartType)
///     }
/// }</code>
/// <para>Creates Services.xlsx with a themed, filterable table of the local services and saves it when the block ends.</para>
/// </example>
/// <example>
/// <summary>Build a regional sales workbook</summary>
/// <code>$orders = Import-OfficeCsv -Path .\Data\Q3-Orders.csv
/// $byRegion = $orders | Group-Object Region | ForEach-Object {
///     [pscustomobject]@{
///         Region  = $_.Name
///         Orders  = $_.Count
///         Revenue = ($_.Group | Measure-Object Amount -Sum).Sum
///     }
/// }
///
/// $properties = @{ Title = 'Q3 sales'; Company = 'Northwind Traders' }
/// New-OfficeExcel -Path .\Reports\Q3-Sales.xlsx -Theme Boardroom -DocumentProperty $properties -Force -Show {
///     ExcelSheet 'Summary' {
///         ExcelTable $byRegion -Name RegionSummary
///         ExcelChart -Table RegionSummary -Address F2 -Type Column -Title 'Revenue by region'
///     }
///     ExcelSheet 'Orders' {
///         ExcelTable $orders -Name Orders
///     }
/// }</code>
/// <para>Builds a Q3 sales workbook in the Boardroom theme with a summary chart and the order detail, sets the document title and opens it in Excel.</para>
/// </example>
/// <example>
/// <summary>Keep a workbook open and save it later</summary>
/// <code>$book = New-OfficeExcel -Theme Modern
/// foreach ($server in 'SQL01', 'SQL02', 'WEB01') {
///     $disks = Invoke-Command -ComputerName $server { Get-Volume | Where-Object DriveLetter }
///     Add-OfficeExcelSheet -Document $book -Name $server {
///         ExcelTable ($disks | Select-Object DriveLetter, FileSystemLabel, Size, SizeRemaining)
///     }
/// }
/// Save-OfficeExcel -Document $book -Path .\Reports\Disk-Capacity.xlsx
/// Close-OfficeExcel -Document $book</code>
/// <para>Without a content block the live workbook is returned, so a script can add a sheet per server in a loop and choose the file name when it saves.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcel", SupportsShouldProcess = true, DefaultParameterSetName = "Live")]
[OutputType(typeof(ExcelDocument), typeof(FileInfo))]
public sealed class NewOfficeExcelCommand : OfficeNewLifecycleCmdlet<ExcelDocument> {
    /// <summary>Built-in theme, theme object or a document/template whose visual theme is read.</summary>
    [Parameter()]
    [OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Existing workbook copied as starting content. Its styling is preserved unless Theme is supplied.</summary>
    [Parameter()]
    [ValidateNotNullOrEmpty]
    public string? Template { get; set; }

    /// <summary>Allow a completed content block to replace an existing destination.</summary>
    [Parameter()]
    public SwitchParameter Force { get; set; }

    /// <summary>Password for a completed content block's saved file; use Save-OfficeExcel for a live workbook.</summary>
    [Parameter(ParameterSetName = "Content")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Native save options or a hashtable, applied to a completed content block.</summary>
    [Parameter(ParameterSetName = "Content")]
    [OfficeExcelSaveOptionsTransformation, ValidateNotNull]
    public ExcelSaveOptions? SaveOptions { get; set; }

    /// <summary>Workbook properties such as Title, Author, Subject, Keywords, Company and Manager.</summary>
    [Parameter()]
    [ValidateNotNull]
    public Hashtable? DocumentProperty { get; set; }

    /// <summary>Native workbook date system.</summary>
    [Parameter()]
    public ExcelDateSystem? DateSystem { get; set; }

    /// <summary>Maximum parallel workers; 1 runs sequentially. Omission retains the native execution policy.</summary>
    [Parameter()]
    [ValidateRange(1, int.MaxValue)]
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
