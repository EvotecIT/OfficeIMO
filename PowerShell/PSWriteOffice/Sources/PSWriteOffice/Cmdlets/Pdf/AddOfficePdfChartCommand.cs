using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a themed chart from data.</summary>
/// <example><summary>Use the native OfficePdfChart command.</summary><code>$sales = @(
///     [pscustomobject]@{ Region = 'North'; Q1 = 1.42; Q2 = 1.55; Q3 = 1.71 }
///     [pscustomobject]@{ Region = 'South'; Q1 = 0.98; Q2 = 1.04; Q3 = 1.12 }
///     [pscustomobject]@{ Region = 'East'; Q1 = 1.21; Q2 = 1.18; Q3 = 1.36 }
///     [pscustomobject]@{ Region = 'West'; Q1 = 1.65; Q2 = 1.79; Q3 = 1.88 }
/// )
/// New-OfficePdf -Path .\Sales.pdf {
///     PdfChart -InputObject $sales -Category Region -Value Q3 -Title 'Q3 revenue by region (USD m)'
/// }</code></example>
[Cmdlet("Add", "OfficePdfChart", DefaultParameterSetName = "Context")]
[Alias("PdfChart") ]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfChartCommand : OfficeCmdlet {
    /// <summary>InputObject for this operation.</summary>
    [Parameter(ParameterSetName = "Context", Mandatory = true, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public object[]? InputObject { get; set; }
    /// <summary>Category for this operation.</summary>
    [Parameter(ParameterSetName = "Context", Mandatory = true)]
    [Parameter(ParameterSetName = "Document", Mandatory = true)]
    public string? Category { get; set; }
    /// <summary>Value for this operation.</summary>
    [Parameter(ParameterSetName = "Context", Mandatory = true)]
    [Parameter(ParameterSetName = "Document", Mandatory = true)]
    public string[]? Value { get; set; }
    /// <summary>Type for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public OfficeChartType Type { get; set; }
    /// <summary>Title for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public string? Title { get; set; }
    /// <summary>Width for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Height for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Alignment for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>AlternativeText for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public string? AlternativeText { get; set; }
    /// <summary>Document for this operation.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }
    /// <summary>Document set only; inside a block the document is returned by New-OfficePdf.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PassThru { get; set; }

    private readonly List<object> _rows = new();
    private bool _collectRows;

    /// <inheritdoc />
    protected override void BeginProcessing() =>
        _collectRows = MyInvocation.ExpectingInput && !MyInvocation.BoundParameters.ContainsKey(nameof(InputObject));

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (_collectRows) {
            _rows.AddRange(InputObject ?? Array.Empty<object>());
            return;
        }
        ComposeChart(InputObject!);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (_collectRows && _rows.Count > 0) ComposeChart(_rows.ToArray());
    }

    private void ComposeChart(object[] rows) {
        var data = PSWriteOffice.Services.OfficeChartDataBuilder.Build(rows, Category!, Value!, Type);
        var snapshot = new OfficeChartSnapshot(Title ?? "Chart", Title, Type, data, Width ?? OfficeLength.Points(450), Height ?? OfficeLength.Points(270));
        var drawing = OfficeChartDrawingRenderer.Render(snapshot);
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, "Document",
            content => content.Drawing(drawing, Alignment, style: new PdfDrawingStyle { AlternativeText = AlternativeText }));
        if (PassThru && document != null) WriteObject(document);
    }
}
