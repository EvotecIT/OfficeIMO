using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

public sealed partial class ConvertOfficeDocumentCommand {
    /// <summary>Caller-owned live document. Conversion captures current edits without saving or closing its source.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    [ValidateNotNull]
    public object Document { get; set; } = null!;

    private async Task ConvertLiveDocumentAsync() {
        if (NoClobber && AutoRename) throw new PSArgumentException("NoClobber and AutoRename cannot be combined.");
        var source = OfficeLiveDocumentSource.Create(Document);
        var target = OfficeFormatSelection.Parse(To)!;
        bool outputText = AsString.IsPresent || OutputPath == null;
        if (outputText && !IsTextExtension(target.Extension)) throw new PSArgumentException("A binary target requires OutputPath.");
        if (AsString && OutputPath != null) throw new PSArgumentException("AsString and OutputPath cannot be combined.");
        if (OfficeFormatSelection.Parse(From) is { } from && !from.Matches(source.Extension))
            throw new PSArgumentException("From must match the live document's native format.");
        string? outputPath = OutputPath == null ? null : OfficePath.ResolveOutput(SessionState, OutputPath, directory: false);
        if (outputPath != null && !target.Matches(System.IO.Path.GetExtension(outputPath)))
            throw new PSArgumentException("To and the OutputPath extension must name the same target format.");
        if (outputPath != null && source.Path != null && OfficePath.AreEquivalent(source.Path, outputPath))
            throw new PSArgumentException("The conversion destination must differ from the live document's associated source.");
        var options = CaptureOptions(new[] { source.Name }, target.Extension);
        options.SourceName = source.Name;
        options.TargetName = outputPath == null ? "output" + target.Extension : System.IO.Path.GetFileName(outputPath);
        if (source.Path != null) options.ResourceDirectory = System.IO.Path.GetDirectoryName(source.Path);
        if (!ShouldProcess(outputPath ?? "in-memory " + target.Extension + " output", "Convert live document")) return;
        byte[] bytes = source.Capture(CancelToken);
        if (bytes.LongLength > options.MaximumInputBytes) throw new PSArgumentException("The live document exceeds MaximumInputBytes.");
        using var input = new MemoryStream(bytes, writable: false);
        OfficeConversionReport report;
        if (outputText) {
            using var output = new MemoryStream();
            report = await OfficeConverter.ConvertAsync(input, OfficeConverter.FormatForExtension(source.Extension), output,
                OfficeConverter.FormatForExtension(target.Extension), options, CancelToken);
            SetSummary(report);
            WriteObject(TextEncoding.GetString(output.ToArray()));
        } else {
            report = await OfficeConverter.ConvertAsync(input, OfficeConverter.FormatForExtension(source.Extension), outputPath!, options, CancelToken);
            SetSummary(report);
            WriteObject(report);
        }
    }
}
