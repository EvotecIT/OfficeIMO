using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;
#if !FRAMEWORK
using OfficeIMO.Invoicing.Validation;
#endif
using System.Threading.Tasks;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Validates an invoice model or exact XML bytes against caller-supplied authority rules.</summary>
/// <para>Model validation reports semantic completeness and arithmetic. File validation requires PowerShell 7.4 and local pinned rule archives; business rules require a caller-installed Saxon/Java runtime. Missing rule execution remains visible in the report.</para>
/// <example>
///   <summary>Reports semantic completeness and arithmetic; this does not run external standards rules.</summary>
///   <code>$invoice | Test-OfficeInvoice</code>
/// </example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeInvoice", DefaultParameterSetName = "Model")]
[OutputType(typeof(InvoiceModelValidationResult))]
#if !FRAMEWORK
[OutputType(typeof(InvoiceValidationReport))]
#endif
public sealed class TestOfficeInvoiceCommand : AsyncPSCmdlet {
    /// <summary>Editable invoice.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Model")]
    public Invoice Invoice { get; set; } = null!;

    /// <summary>Exact invoice XML bytes to validate against pinned local authority rules. Requires PowerShell 7.4.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "File")]
    [Alias("FullName", "LiteralPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Pinned KoSIT rule archive. No rules are downloaded automatically.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "File")]
    public string RulesArchivePath { get; set; } = string.Empty;

    /// <summary>Optional pinned Factur-X validation archive.</summary>
    [Parameter(ParameterSetName = "File")]
    public string? FacturXArchivePath { get; set; }

    /// <summary>Optional pinned Peppol Schematron file.</summary>
    [Parameter(ParameterSetName = "File")]
    public string? PeppolRulesPath { get; set; }

    /// <summary>Exact specification for file validation.</summary>
    [Parameter(ParameterSetName = "File")]
    public InvoiceSpecificationRelease SpecificationRelease { get; set; } = InvoiceSpecificationRelease.En16931_1_3_16;

    /// <summary>Optional caller-installed Saxon HE JAR for business rules. Without it, business rules remain NotRun and IsValid is false.</summary>
    [Parameter(ParameterSetName = "File")]
    public string? SaxonJarPath { get; set; }

    /// <summary>Caller-installed Java executable used only when SaxonJarPath is supplied.</summary>
    [Parameter(ParameterSetName = "File")]
    public string JavaExecutable { get; set; } = "java";

    /// <summary>Maximum business-rule process duration in seconds.</summary>
    [Parameter(ParameterSetName = "File")]
    [ValidateRange(1, 3600)]
    public int TimeoutSeconds { get; set; } = 120;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ParameterSetName == "Model") {
            WriteObject(InvoiceCommandUtilities.Validate(Invoice));
            return;
        }
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Authority-rule invoice validation requires PowerShell 7.4 or newer. Model validation is available in this host.");
#else
        string Resolve(string path) => OfficeFileCommandUtilities.ResolvePath(this, path);
        var bundle = InvoiceRuleBundle.Load(Resolve(RulesArchivePath),
            PeppolRulesPath == null ? null : Resolve(PeppolRulesPath), FacturXArchivePath == null ? null : Resolve(FacturXArchivePath));
        var runner = SaxonJarPath == null ? null : new SaxonInvoiceRulesRunner(Resolve(SaxonJarPath), JavaExecutable, TimeSpan.FromSeconds(TimeoutSeconds));
        using var stream = File.OpenRead(Resolve(Path));
        var bytes = OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(stream, InvoiceProfileDeclaration.MaximumXmlBytes);
        WriteObject(await new InvoiceValidator(bundle, runner).ValidateAsync(bytes, SpecificationRelease, CancelToken));
#endif
    }
}
