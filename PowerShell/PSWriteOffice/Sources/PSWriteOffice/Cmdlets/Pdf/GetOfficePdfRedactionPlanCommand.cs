using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Previews explicit, literal-text or bounded regular-expression redaction selections without changing PDF content.</summary>
/// <example><summary>Preview sensitive identifiers.</summary><code>Get-OfficePdfRedactionPlan -LiteralPath './Report[1].pdf' -Pattern 'Account: \d+'</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePdfRedactionPlan", DefaultParameterSetName = "Text")]
[OutputType(typeof(PdfRedactionPlan))]
public sealed partial class GetOfficePdfRedactionPlanCommand : PSCmdlet {
    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool literal = ParameterSetName.EndsWith("LiteralPath", StringComparison.Ordinal);
        var options = PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions.IsPresent);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(IgnorePermissionRestrictions)) is bool ignore)
            options = (options ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        foreach (string input in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
            using var document = PdfDocument.Load(input, options);
            if (ParameterSetName.StartsWith("Area", StringComparison.Ordinal)) WriteObject(document.Redactions.Plan(Area));
            else if (ParameterSetName.StartsWith("Rectangle", StringComparison.Ordinal))
                WriteObject(document.Redactions.Plan(new[] { new PdfRedactionArea(Page, X.ToPoints(), Y.ToPoints(), Width.ToPoints(), Height.ToPoints(), Label) }));
            else {
                bool regex = ParameterSetName.StartsWith("Pattern", StringComparison.Ordinal);
                var search = PdfRedactionCommandOptions.Create(regex ? null : Text, regex ? Pattern : null, CaseSensitive.IsPresent);
                WriteObject(document.Redactions.Search(search));
            }
        }
    }
}
