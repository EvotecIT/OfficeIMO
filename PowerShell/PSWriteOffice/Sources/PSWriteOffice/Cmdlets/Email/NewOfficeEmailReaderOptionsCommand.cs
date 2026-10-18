using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Creates bounded EML, MSG, and TNEF reader settings through ordinary PowerShell parameters.</summary>
/// <example>
///   <summary>Read message diagnostics without retaining attachment payloads.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeEmailReaderOptions -ExcludeAttachmentContent -MaxAttachmentBytes 25MB
/// Get-OfficeEmail -Path .\Message.msg -Options $options -AsResult</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeEmailReaderOptions")]
[OutputType(typeof(EmailReaderOptions))]
public sealed class NewOfficeEmailReaderOptionsCommand : PSCmdlet {
    /// <summary>Maximum artifact size accepted by the reader.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxInputBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Maximum bytes allowed in one MIME header section.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxHeaderBytes { get; set; } = 1024 * 1024;
    /// <summary>Maximum number of header fields in one entity.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxHeaderCount { get; set; } = 10000;
    /// <summary>Maximum MIME entity count.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxPartCount { get; set; } = 10000;
    /// <summary>Maximum nested MIME depth.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxMimeDepth { get; set; } = 64;
    /// <summary>Maximum decoded bytes for one attachment.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxAttachmentBytes { get; set; } = 128L * 1024L * 1024L;
    /// <summary>Maximum aggregate decoded attachment bytes.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxTotalAttachmentBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Maximum embedded-message recursion depth.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int MaxNestedMessageDepth { get; set; } = 16;
    /// <summary>Do not retain decoded attachment payloads in memory.</summary>
    [Parameter] public SwitchParameter ExcludeAttachmentContent { get; set; }
    /// <summary>Retain original artifact bytes for an explicit lossless write.</summary>
    [Parameter] public SwitchParameter PreserveRawSource { get; set; }
    /// <summary>Maximum CFB directory entries accepted while reading MSG.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxCompoundDirectoryEntries { get; set; } = 65536;
    /// <summary>Maximum aggregate MAPI properties across a message and embedded messages.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxMapiPropertyCount { get; set; } = 100000;
    /// <summary>Maximum aggregate bytes represented by decoded MSG property streams.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxDecodedPropertyBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Maximum number of TNEF attributes.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxTnefAttributeCount { get; set; } = 100000;
    /// <summary>Maximum aggregate attachment count.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxAttachmentCount { get; set; } = 10000;

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(new EmailReaderOptions(
        MaxInputBytes,
        MaxHeaderBytes,
        MaxHeaderCount,
        MaxPartCount,
        MaxMimeDepth,
        MaxAttachmentBytes,
        MaxTotalAttachmentBytes,
        MaxNestedMessageDepth,
        !ExcludeAttachmentContent.IsPresent,
        PreserveRawSource.IsPresent,
        MaxCompoundDirectoryEntries,
        MaxMapiPropertyCount,
        MaxDecodedPropertyBytes,
        MaxTnefAttributeCount,
        MaxAttachmentCount));
}
