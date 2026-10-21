using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;
using OfficeIMO.Core.Internal;
namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Saves CII or UBL with strict source and target loss handling.</summary>
/// <example>
///   <summary>Saves CII or UBL with strict source and target loss handling.</summary>
///   <code>$invoice | Save-OfficeInvoice -Path ./invoice.xml -Syntax Ubl</code>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeInvoice", SupportsShouldProcess = true)]
[OutputType(typeof(Invoice))]
public sealed class SaveOfficeInvoiceCommand : PSCmdlet {
    /// <summary>Editable invoice.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public Invoice Invoice { get; set; } = null!;

    /// <summary>Destination XML filename.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Exact supported output specification.</summary>
    [Parameter]
    public InvoiceSpecificationRelease SpecificationRelease { get; set; } = InvoiceSpecificationRelease.En16931_1_3_16;

    /// <summary>Output XML syntax.</summary>
    [Parameter]
    public InvoiceSyntax Syntax { get; set; } = InvoiceSyntax.Ubl;

    /// <summary>Output profile; supported combinations are listed by Get-OfficeInvoiceTarget.</summary>
    [Parameter]
    public InvoiceProfile Profile { get; set; } = InvoiceProfile.En16931;

    /// <summary>Permit only profile-defined reductions supported by the target engine.</summary>
    [Parameter]
    public SwitchParameter AllowProfileLoss { get; set; }

    /// <summary>Explicitly permit discarding unmapped source business data.</summary>
    [Parameter]
    public SwitchParameter AllowUnmappedDataLoss { get; set; }

    /// <summary>Replace an existing output.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Variable receiving target-mapping diagnostics.</summary>
    [Parameter]
    public string? ReportVariable { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var path = OfficeFileCommandUtilities.ResolvePath(this, Path);
        if (!ShouldProcess(path, "Save electronic invoice XML")) return;
        var target = new InvoiceXmlOptions(SpecificationRelease, Syntax, Profile,
            AllowProfileLoss.IsPresent ? InvoiceProjectionPolicy.AllowProfileDefinedDataLoss : InvoiceProjectionPolicy.RejectDataLoss);
        var diagnostics = InvoiceSerializer.InspectTarget(Invoice, target);
        byte[] bytes = InvoiceCommandUtilities.Write(Invoice, target, AllowUnmappedDataLoss.IsPresent);
        OfficeFileOutput.Write(path, bytes, Force.IsPresent);
        if (!string.IsNullOrWhiteSpace(ReportVariable)) SessionState.PSVariable.Set(ReportVariable, diagnostics);
        if (PassThru.IsPresent) WriteObject(Invoice);
    }
}
