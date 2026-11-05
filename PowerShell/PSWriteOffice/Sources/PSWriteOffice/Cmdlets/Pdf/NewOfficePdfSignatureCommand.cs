using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Threading;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Prepare external PDF signature using native byte-range and signature-placeholder operations.</summary>
/// <para>Confirmation precedes loading, native mutation and output creation. This command does not obtain certificates or private keys.</para>
[Cmdlet(VerbsCommon.New, "OfficePdfSignature", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(PdfExternalSignaturePreparation))]
public sealed partial class NewOfficePdfSignatureCommand : PSCmdlet, IDisposable {
    private readonly List<string> _inputs = new();
    private readonly CancellationTokenSource _stop = new();
    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
        _inputs.AddRange(OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal));
    }
    /// <inheritdoc />
    protected override void EndProcessing() {
        var read = PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions.IsPresent);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(IgnorePermissionRestrictions)) is bool ignore)
            read = (read ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        foreach (var item in PdfFileCommandInputs.Resolve(this, null, _inputs.ToArray(), OutputPath, OutputDirectory, false)) {
            if (!ShouldProcess(item.Output, "Prepare external PDF signature")) continue;
            using var document = PdfDocument.Load(item.Input, read, _stop.Token);
            var options = new PdfExternalSignatureOptions { FieldName = FieldName!, Filter = Filter!, SubFilter = SubFilter,
                Name = Name, Reason = Reason, Location = Location, ContactInfo = ContactInfo, ReservedSignatureContentsBytes = ReservedSize };
            var preparation = document.Security.PrepareExternalSignature(options, _stop.Token);
            _stop.Token.ThrowIfCancellationRequested();
            PdfCommandUtilities.EnsureDirectory(item.Output);
            OfficeFileOutput.Write(item.Output, true, temporary => File.WriteAllBytes(temporary, preparation.PreparedPdf));
            WriteObject(PassThru ? preparation : new FileInfo(item.Output));
        }
    }
    /// <inheritdoc />
    protected override void StopProcessing() => _stop.Cancel();
    /// <summary>Releases the command-owned cancellation source.</summary>
    public void Dispose() => _stop.Dispose();
}
