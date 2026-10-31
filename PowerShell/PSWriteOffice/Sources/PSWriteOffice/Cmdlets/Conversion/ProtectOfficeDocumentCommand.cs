using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Adds native password encryption or editing restrictions to a document.</summary>
/// <para>Supports PDF and Open XML Word, Excel and PowerPoint packages. Path input is saved in place unless
/// OutputPath names a protected copy. File inputs must be unencrypted: this command has no CurrentPassword
/// parameter. Live Word/Excel restrictions may remain in memory; encryption requires OutputPath or a native
/// associated file path. Encryption is a one-time export, not a stored password setting on the live input.
/// PDF protection returns a new native document with preservation proof, leaving the original live PDF intact.
/// Without Encrypt, PDF restrictions use Password or RestrictionPassword as the owner password and allow opening
/// without a password. With Encrypt, a separate owner password keeps the open password from authorizing permission
/// changes. Unsupported restrictions produce warnings; an entirely unsupported request writes nothing.
/// Safe native signature and preservation policies remain in force.</para>
/// <example>
///   <summary>Encrypt a workbook using a secure password.</summary>
///   <code>Get-ChildItem ./Payroll -Filter *.xlsx | Protect-OfficeDocument -Encrypt -Password (Read-Host -AsSecureString) -PassThru</code>
/// </example>
/// <example>
///   <summary>Restrict Word editing without encrypting the file.</summary>
///   <code>Protect-OfficeDocument -Path ./policy.docx -Restriction ReadOnly -Password (Read-Host -AsSecureString)</code>
/// </example>
/// <example>
///   <summary>Withhold PDF copying with separate open and owner passwords.</summary>
///   <code>Protect-OfficeDocument -Path ./board.pdf -Encrypt -Password $openPassword -Restriction Copy -RestrictionPassword $ownerPassword -OutputPath ./protected.pdf</code>
/// </example>
/// <example>
///   <summary>Withhold PDF copying while allowing readers to open without a password.</summary>
///   <code>Protect-OfficeDocument -Path ./board.pdf -Restriction Copy -Password $ownerPassword -OutputPath ./restricted.pdf</code>
/// </example>
[Cmdlet(VerbsSecurity.Protect, "OfficeDocument", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(OfficeIMO.Word.WordDocument), typeof(OfficeIMO.Excel.ExcelDocument),
    typeof(OfficeIMO.PowerPoint.PowerPointDocument), typeof(PdfDocument))]
public sealed class ProtectOfficeDocumentCommand : PSCmdlet {
    private readonly List<string> _paths = new();
    private readonly List<object> _documents = new();

    /// <summary>Files to protect, in pipeline and provider match order.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal files, including FileInfo.FullName and provider PSPath properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned native Word, Excel, PowerPoint or PDF document.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    [ValidateNotNull]
    public IDisposable Document { get; set; } = null!;

    /// <summary>Open password with Encrypt; otherwise the restriction password. Accepts a credential or warned plain text.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Encrypt the saved artifact using its native engine.</summary>
    [Parameter]
    public SwitchParameter Encrypt { get; set; }

    /// <summary>Word editing mode, Excel protection/reservation, or PDF permissions to withhold.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public OfficeProtectionRestriction[] Restriction { get; set; } = Array.Empty<OfficeProtectionRestriction>();

    /// <summary>Restriction password, separate from the open password when Encrypt is used; PDF owner password.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? RestrictionPassword { get; set; }

    /// <summary>Protected file destination; accepts exactly one input when supplied.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }

    /// <summary>Native Excel protection/reservation or PDF encryption options, with property hashtable overrides.</summary>
    /// <para>Excel tables accept LegacyPasswordHash, ModifyPasswordHash and ReservedBy; native workbook and reservation
    /// property tables are also accepted. PDF tables accept Algorithm, EncryptMetadata, Permissions or AllowedPermissions,
    /// and AesCryptographyProvider. Supply PDF passwords through Password and RestrictionPassword. Typed native PDF
    /// options retain their owner password, permission mask, algorithm, metadata setting and AES provider.</para>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public object[]? Options { get; set; }

    /// <summary>Emit FileInfo for file input, or the protected native document for Document input.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            _documents.Add(OfficeProtectionOptionBinding.Unwrap(Document) ?? throw new PSArgumentException("Document cannot be null."));
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string path in literal ? LiteralPath : Path)
            _paths.AddRange(OfficePath.ResolveInput(SessionState, path, literal));
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        bool live = ParameterSetName == "Document";
        int count = live ? _documents.Count : _paths.Count;
        if (count == 0) throw new PSArgumentException("Provide at least one input document.");
        if (OutputPath != null && count != 1) throw new PSArgumentException("OutputPath requires exactly one input document.");
        foreach (var restriction in Restriction)
            if (!Enum.IsDefined(typeof(OfficeProtectionRestriction), restriction)) throw new PSArgumentException("Invalid Restriction value.");
        string? password = Plaintext(Password), restrictionPassword = Plaintext(RestrictionPassword);
        if (Password != null && string.IsNullOrWhiteSpace(password)) throw new PSArgumentException("Password cannot be empty or whitespace.");
        if (RestrictionPassword != null && string.IsNullOrWhiteSpace(restrictionPassword)) throw new PSArgumentException("RestrictionPassword cannot be empty or whitespace.");
        if (Encrypt.IsPresent && password == null) throw new PSArgumentException("Encrypt requires a non-empty Password.");
        if (!Encrypt.IsPresent && password != null && restrictionPassword != null)
            throw new PSArgumentException("Without Encrypt, supply either Password or RestrictionPassword for restrictions.");
        if (!Encrypt.IsPresent) restrictionPassword ??= password;
        if (!Encrypt.IsPresent && Restriction.Length == 0 && Options == null)
            throw new PSArgumentException("Specify Encrypt, Restriction, or native protection Options.");

        string? explicitOutput = OutputPath == null ? null : OfficePath.ResolveOutput(SessionState, OutputPath, directory: false);
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        var plans = new List<ProtectionPlan>();
        var owned = new List<IDisposable>();
        Exception? failure = null;
        try {
            // Capture and validate the entire batch before changing any live document or destination.
            for (int index = 0; index < count; index++) {
                object? document = live ? _documents[index] : null;
                string extension = document == null ? System.IO.Path.GetExtension(_paths[index]) : OfficeProtectionDispatch.Extension(document);
                var family = OfficeProtectionDispatch.Family(extension);
                var supported = Restriction.Distinct().Where(value => OfficeProtectionDispatch.Supports(family, value)).ToArray();
                foreach (var unsupported in Restriction.Distinct().Except(supported))
                    WriteWarning($"{family} does not support restriction {unsupported}; that restriction is not applied.");
                if (family == OfficeProtectionFamily.Word && supported.Length > 1)
                    throw new PSArgumentException("Word supports exactly one editing restriction at a time.");
                bool permissionsOnly = family == OfficeProtectionFamily.Pdf && !Encrypt.IsPresent && supported.Length > 0;
                if (permissionsOnly && restrictionPassword == null)
                    throw new PSArgumentException("PDF permissions-only protection requires Password or RestrictionPassword as the owner password.");
                var options = OfficeProtectionOptionBinding.Capture(engine, family, Options,
                    Encrypt.IsPresent ? password : null, permissionsOnly ? restrictionPassword : null);
                if (family == OfficeProtectionFamily.Excel) options.PrepareExcel(supported, restrictionPassword);
                bool changes = Encrypt.IsPresent || supported.Length > 0 || options.HasWorkbook || options.HasReservation;
                if (!changes && Restriction.Length == 0)
                    throw new PSArgumentException("The supplied Options do not select any native protection operation.");
                if (family == OfficeProtectionFamily.Word && supported.Length > 0 && restrictionPassword == null)
                    throw new PSArgumentException("Word editing restrictions require Password or RestrictionPassword.");
                if (family == OfficeProtectionFamily.Pdf && changes) {
                    if (restrictionPassword != null) options.Pdf!.OwnerPassword = restrictionPassword;
                    string? owner = restrictionPassword ?? options.Pdf?.OwnerPassword;
                    bool withholds = supported.Length > 0 || options.Pdf!.AllowedPermissions != PdfStandardPermissions.All;
                    if (password?.Length > 4096 || owner?.Length > 4096)
                        throw new PSArgumentException("PDF passwords must not exceed the native 4096-character input limit.");
                    if (Encrypt.IsPresent && withholds && !options.Pdf!.HasDistinctOwnerPassword)
                        WriteWarning("The native PDF options reuse the open password for owner authorization. Supply a separate RestrictionPassword to keep open-password holders from bypassing permissions.");
                }
                string? associated = document == null ? null : OfficeProtectionDispatch.AssociatedPath(document);
                string? output = explicitOutput ?? (live ? Encrypt.IsPresent ? associated : null : _paths[index]);
                if (Encrypt.IsPresent && string.IsNullOrWhiteSpace(output))
                    throw new PSArgumentException("Live encryption requires OutputPath or an actual native associated file path.");
                if (output != null) {
                    output = OfficePath.ResolveOutput(SessionState, output, directory: false);
                    if (!extension.Equals(System.IO.Path.GetExtension(output), StringComparison.OrdinalIgnoreCase))
                        throw new PSArgumentException("OutputPath must have the same native filename format and extension as the input.");
                    OfficePath.ValidateDestination(output, refuseExisting: false);
                }
                if (!changes) continue; // Unsupported-only restrictions never create a no-op copy.
                if (document == null) {
                    var loaded = OfficeProtectionDispatch.Load(_paths[index], family);
                    owned.Add(loaded);
                    document = loaded;
                }
                OfficeProtectionDispatch.Validate(document, extension, changes, Encrypt.IsPresent, output);
                plans.Add(new ProtectionPlan(document, supported, options, output));
            }
            foreach (var plan in plans) {
                if (!ShouldProcess(plan.Output ?? "Live " + plan.Document.GetType().Name, "Protect document")) continue;
                object result = OfficeProtectionDispatch.Apply(plan.Document, plan.Restrictions, plan.Options,
                    Encrypt.IsPresent, password, restrictionPassword, plan.Output, WriteVerbose);
                bool handedToCaller = false;
                try {
                    if (PassThru.IsPresent) {
                        WriteObject(live ? result : new FileInfo(plan.Output!));
                        handedToCaller = live;
                    }
                } finally {
                    // A new PDF stays owned here until successful live PassThru, including a pipeline stop
                    // or an exception while emitting a saved FileInfo. Original inputs retain their own lifetime.
                    if (!ReferenceEquals(result, plan.Document) && !handedToCaller) ((IDisposable)result).Dispose();
                }
            }
        } catch (Exception error) { failure = error; throw; }
        finally {
            var disposalErrors = new List<Exception>();
            foreach (IDisposable document in owned) {
                try { document.Dispose(); } catch (Exception error) { disposalErrors.Add(error); }
            }
            if (failure == null && disposalErrors.Count > 0) throw new AggregateException(disposalErrors);
        }
    }

    private static string? Plaintext(SecureString? value) => value == null ? null : new NetworkCredential(string.Empty, value).Password;

    private sealed class ProtectionPlan {
        internal ProtectionPlan(object document, OfficeProtectionRestriction[] restrictions, OfficeProtectionOptionBinding options, string? output) {
            Document = document; Restrictions = restrictions; Options = options; Output = output;
        }
        internal object Document { get; }
        internal OfficeProtectionRestriction[] Restrictions { get; }
        internal OfficeProtectionOptionBinding Options { get; }
        internal string? Output { get; }
    }
}
