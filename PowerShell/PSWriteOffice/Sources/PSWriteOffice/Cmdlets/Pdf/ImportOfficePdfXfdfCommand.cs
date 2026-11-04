using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Threading;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Imports bounded DTD-free XFDF into native PDF forms, with atomic file output or a caller-owned live result.</summary>
/// <example><summary>Fill a literal file from XFDF.</summary><code>Import-OfficePdfXfdf -LiteralPath './Form[1].pdf' -XfdfPath ./Values.xfdf -OutputPath ./Completed.pdf</code></example>
[Cmdlet(VerbsData.Import, "OfficePdfXfdf", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument))]
public sealed partial class ImportOfficePdfXfdfCommand : OfficeCmdlet, IDisposable {
    private readonly StringBuilder _xml = new();
    private readonly List<string> _inputs = new();
    private readonly List<PdfDocument> _documents = new();
    private readonly CancellationTokenSource _stop = new();
    private long _bytes;
    private bool _hasText;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (MaxXfdfSize < 1) throw new PSArgumentOutOfRangeException(nameof(MaxXfdfSize));
        bool file = ParameterSetName.Contains("File");
        if (!file) {
            long additional = Encoding.UTF8.GetByteCount(Xfdf ?? string.Empty) + (_hasText ? 1L : 0L);
            if (additional > MaxXfdfSize - _bytes) throw new InvalidDataException("XFDF input exceeds MaxXfdfSize.");
            if (_hasText) _xml.Append('\n');
            _xml.Append(Xfdf); _bytes += additional; _hasText = true;
        }
        if (ParameterSetName.StartsWith("Document", StringComparison.Ordinal)) {
            var document = Document ?? OfficeDocumentContext.Find<PdfDocument>() ?? throw new PSArgumentException("Provide Document or use an Open-/New-OfficePdf block.");
            if (file || _documents.Count == 0) _documents.Add(document);
        } else if (file || _inputs.Count == 0) {
            bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
            _inputs.AddRange(OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal));
        }
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        var binder = new OfficeOptionsTransformationAttribute(typeof(PdfFormFillerOptions));
        PdfFormFillerOptions options = new();
        foreach (object item in Options ?? Array.Empty<object>()) {
            object? value = item is PSObject wrapped ? wrapped.BaseObject : item;
            if (value is PdfFormFillerOptions native) options = native.Clone();
            else if (value is IDictionary fields) options = (PdfFormFillerOptions)binder.ApplyTo((EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext"), options, fields);
            else throw new PSArgumentException("XFDF Options accepts PdfFormFillerOptions or its property hashtables.");
        }
        string ReadXml() {
            if (!ParameterSetName.Contains("File")) return _xml.ToString();
            string path = OfficePath.ResolveOutput(SessionState, XfdfPath!, false);
            if (new FileInfo(path).Length > MaxXfdfSize) throw new InvalidDataException("XFDF file exceeds MaxXfdfSize.");
            string xml = File.ReadAllText(path);
            if (Encoding.UTF8.GetByteCount(xml) > MaxXfdfSize) throw new InvalidDataException("Decoded XFDF exceeds MaxXfdfSize.");
            return xml;
        }
        foreach (var document in _documents) {
            if (!ShouldProcess("PDF document", "Import XFDF form data")) continue;
            var changed = document.Forms.ImportXfdf(ReadXml(), options, _stop.Token);
            bool adopted = OfficeDocumentContext.Replace(document, changed);
            if (PassThru || !adopted) WriteObject(changed);
        }
        var read = PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions.IsPresent);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(IgnorePermissionRestrictions)) is bool ignore)
            read = (read ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        if (_inputs.Count == 0) return;
        foreach (var item in PdfFileCommandInputs.Resolve(this, null, _inputs.ToArray(), OutputPath, OutputDirectory, false)) {
            if (!ShouldProcess(item.Output, "Import XFDF form data")) continue;
            using var document = PdfDocument.Load(item.Input, read, _stop.Token);
            var changed = document.Forms.ImportXfdf(ReadXml(), options, _stop.Token);
            bool transferred = false;
            try {
                changed.Save(item.Output, new PdfSaveOptions { ConflictPolicy = OfficeConversionFileConflictPolicy.Replace }, _stop.Token);
                if (PassThru) {
                    WriteObject(changed);
                    transferred = true;
                }
            } finally { if (!transferred) changed.Dispose(); }
        }
    }
    /// <inheritdoc />
    protected override void StopProcessing() => _stop.Cancel();
    /// <summary>Releases the command-owned cancellation source.</summary>
    public void Dispose() => _stop.Dispose();
}
