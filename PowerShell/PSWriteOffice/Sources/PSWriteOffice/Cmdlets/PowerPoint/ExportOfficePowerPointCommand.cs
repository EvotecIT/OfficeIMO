using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Exports objects as a designed presentation with editable tables and optional charts.</summary>
/// <para>Creates a title slide, an optional grouped summary, repeating table headers and native continuation slides. Replaces an existing file unless NoClobber is supplied. Output is quiet unless PassThru is supplied; publication is atomic.</para>
/// <example>
/// <summary>Present regional revenue with a chart.</summary>
/// <code>$sales | Export-OfficePowerPoint -Path ./Sales.pptx -Title 'Regional sales' -Property Region,Revenue -Category Region -Value Revenue -Theme Boardroom -PassThru</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficePowerPoint", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficePowerPointCommand : OfficeCmdlet {
    private readonly List<object?> _input = new();

    /// <summary>Objects supplied as an array or through the pipeline.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Destination PPTX file; its parent folder is created only when publishing.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Title slide text; omission uses the destination file name without its extension.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Title { get; set; }

    /// <summary>Subtitle on the opening slide.</summary>
    [Parameter]
    public string? Subtitle { get; set; }

    /// <summary>Table properties in display order; omission includes all properties.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string[]? Property { get; set; }

    /// <summary>Property containing chart categories.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Category { get; set; }

    /// <summary>Numeric properties plotted as chart series in the supplied order.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string[]? Value { get; set; }

    /// <summary>Shared native chart kind; defaults to clustered columns.</summary>
    [Parameter]
    public OfficeChartKind ChartType { get; set; } = OfficeChartKind.ColumnClustered;

    /// <summary>Chart slide title; omission uses the category property name.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? ChartTitle { get; set; }

    /// <summary>Property counted on the summary; omission tries Category, then a text property with two to six groups.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? GroupBy { get; set; }

    /// <summary>Omit the At a glance summary.</summary>
    [Parameter]
    public SwitchParameter NoSummary { get; set; }

    /// <summary>Export only title and chart slides; requires Category and Value.</summary>
    [Parameter]
    public SwitchParameter NoTable { get; set; }

    /// <summary>Maximum data rows per physical table slide; omission uses native readable capacity.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxRows { get; set; }

    /// <summary>Built-in theme, OfficeTheme model or PowerPoint template path; defaults to Modern.</summary>
    [Parameter, OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Brand accent, resolved against the selected theme.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor Accent { get; set; }

    /// <summary>Culture for displayed cell values, group counts and numeric chart input; defaults to invariant.</summary>
    [Parameter, ValidateNotNull]
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>Refuse to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Open the saved presentation using the system's default application.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Return the saved FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var item in InputObject)
            TableInputCollector.AddInput(_input, item, preserveTabularInput: true);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        string? path = null;
        try {
            path = ResolveOutputPath(Path);
            if (!System.IO.Path.GetExtension(path).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Export-OfficePowerPoint writes a .pptx presentation.", nameof(Path));
            var plan = PowerPointObjectExport.CreatePlan(_input, Title ?? System.IO.Path.GetFileNameWithoutExtension(path),
                Subtitle, Property, Category, Value, ChartType, ChartTitle, GroupBy,
                NoSummary, NoTable, MaxRows, Culture,
                MyInvocation.BoundParameters.ContainsKey(nameof(ChartType)));
            using var document = PowerPointObjectExport.CreatePresentation(Theme,
                MyInvocation.BoundParameters.ContainsKey(nameof(Accent)) ? Accent : (OfficeColor?)null);
            var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand(
                document.GetThemeColor(PowerPointThemeColor.Accent1) ?? "2563EB", "object-export"));
            options.ApplyTheme = false;
            document.Compose(plan, options);
            if (!ShouldProcess(path, "Export PowerPoint presentation")) return;
            OfficePath.ValidateDestination(path, NoClobber);
            // Strict native serialization completes before the atomic writer can create a folder or temporary file.
            byte[] bytes = document.ToBytes(new PowerPointSaveOptions());
            OfficeOutput.Write(path, stream => stream.Write(bytes, 0, bytes.Length), NoClobber);
            if (Show) OfficeShow.Open(path);
            if (PassThru) WriteObject(new FileInfo(path));
        } catch (ArgumentException error) { InvalidArgument(error, path ?? Path); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-object-export-failed", path ?? Path);
        }
    }
}
