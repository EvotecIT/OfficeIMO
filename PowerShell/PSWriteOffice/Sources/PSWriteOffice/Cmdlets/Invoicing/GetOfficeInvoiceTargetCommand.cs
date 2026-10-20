using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Lists exact supported invoice authoring contracts from the engine.</summary>
/// <example>
///   <summary>Lists exact supported invoice authoring contracts from the engine.</summary>
///   <code>Get-OfficeInvoiceTarget -Syntax Ubl</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeInvoiceTarget")]
[OutputType(typeof(InvoiceXmlOptions))]
public sealed class GetOfficeInvoiceTargetCommand : PSCmdlet {
    /// <summary>Optional XML syntax filter.</summary>
    [Parameter]
    public InvoiceSyntax? Syntax { get; set; }

    /// <summary>Optional guideline profile filter.</summary>
    [Parameter]
    public InvoiceProfile? Profile { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var target in InvoiceXmlOptions.GetSupportedTargets()) {
            if (Syntax.HasValue && target.Syntax != Syntax.Value) continue;
            if (Profile.HasValue && target.Profile != Profile.Value) continue;
            WriteObject(target);
        }
    }
}
