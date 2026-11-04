using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Creates an immutable document reader with common policies, format options and optional OCR.</summary>
/// <example><summary>Reuse one configured reader across a report folder.</summary>
/// <code>$options = New-OfficeDocumentReaderOptions -Projection TextAndMetadata
/// $reader = New-OfficeDocumentReader -Options $options -MaximumConcurrentReads 4
/// Read-OfficeDocument ./Reports -Recurse -Reader $reader -As Text</code></example>
/// <example><summary>Configure optional OCR with one settings object.</summary>
/// <code>$ocr = New-OfficeOcrOptions -Language English,Polish -NoLanguageDownload
/// $reader = New-OfficeDocumentReader -OcrOptions $ocr
/// Read-OfficeDocument ./Scans -Reader $reader</code></example>
[Cmdlet(VerbsCommon.New, "OfficeDocumentReader", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeDocumentReader))]
public sealed class NewOfficeDocumentReaderCommand : AsyncPSCmdlet {
    /// <summary>Typed common and format options, or common/format-keyed hashtables.</summary>
    [Parameter] [OfficeReadOptionsTransformation] public object[]? Options { get; set; }
    /// <summary>Ordered processors applied after native extraction.</summary>
    [Parameter] public IOfficeDocumentProcessor[]? Processor { get; set; }
    /// <summary>How ordinary processor failures affect the operation. Safety limits remain terminal.</summary>
    [Parameter] public OfficeDocumentProcessorFailureBehavior ProcessorFailureBehavior { get; set; } = OfficeDocumentProcessorFailureBehavior.Throw;
    /// <summary>Enable the default local OCR provider.</summary>
    [Parameter] public SwitchParameter Ocr { get; set; }
    /// <summary>Optional OCR provider and execution settings. Implies Ocr.</summary>
    [Parameter] [OfficeOptionsTransformation(typeof(OfficeOcrOptions))] public OfficeOcrOptions? OcrOptions { get; set; }
    /// <summary>Maximum items projected from each mail store; zero means all.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int MaximumStoreItems { get; set; } = 1000;
    /// <summary>Maximum concurrent native document reads.</summary>
    [Parameter] [ValidateRange(1, 64)] public int MaximumConcurrentReads { get; set; } = 4;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        var configuration = OfficeReadConfiguration.Bind(Options);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumStoreItems)) || configuration.Formats.Email?.Stores == null) {
            configuration.Formats.Email = OfficeReadConfiguration.Copy(configuration.Formats.Email);
            configuration.Formats.Email.Stores = OfficeReadConfiguration.Copy(configuration.Formats.Email.Stores);
            configuration.Formats.Email.Stores.MaxItems = MaximumStoreItems == 0 ? int.MaxValue : MaximumStoreItems;
        }
        var builder = ReaderCommandUtilities.CreateBuilder(configuration.Formats)
            .WithMaxConcurrentReads(MaximumConcurrentReads).WithProcessorFailureBehavior(ProcessorFailureBehavior);
        if (configuration.HasCommonOptions) builder.WithDefaultOptions(configuration.Common);
        if (Processor != null) builder.AddProcessors(Processor);
        if (Ocr || OcrOptions != null) {
            var settings = OfficeReadOcr.Capture(OcrOptions ?? new OfficeOcrOptions());
            var engine = await OfficeReadOcr.CreateEngineAsync(settings, CancelToken);
            builder.AddProcessor(new OfficeDocumentOcrProcessor(engine, settings.ExecutionOptions));
        }
        WriteObject(builder.Build());
    }
}
