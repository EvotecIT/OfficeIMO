using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Creates deterministic email writer settings through ordinary PowerShell parameters.</summary>
/// <example>
///   <summary>Preserve the original source when possible and block semantic loss.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeEmailWriterOptions -UsePreservedRawSource -ConversionLossPolicy Block
/// $message | Save-OfficeEmail -Path .\Message.eml -Options $options -PassThru</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeEmailWriterOptions")]
[OutputType(typeof(EmailWriterOptions))]
public sealed class NewOfficeEmailWriterOptionsCommand : PSCmdlet {
    /// <summary>Policy applied when the requested format cannot preserve known message semantics.</summary>
    [Parameter] public EmailConversionLossPolicy ConversionLossPolicy { get; set; } = EmailConversionLossPolicy.Block;
    /// <summary>Emit an unchanged preserved source instead of regenerating the artifact when possible.</summary>
    [Parameter] public SwitchParameter UsePreservedRawSource { get; set; }
    /// <summary>Write Bcc recipients into the message header.</summary>
    [Parameter] public SwitchParameter IncludeBccHeader { get; set; }
    /// <summary>Maximum encoded characters on one Base64 body line.</summary>
    [Parameter] [ValidateRange(4, 996)] public int Base64LineLength { get; set; } = 76;
    /// <summary>Maximum embedded-message write depth.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int MaxNestedMessageDepth { get; set; } = 16;
    /// <summary>Maximum serialized artifact size.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxOutputBytes { get; set; } = 512L * 1024L * 1024L;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Base64LineLength % 4 != 0) {
            throw new PSArgumentOutOfRangeException(nameof(Base64LineLength), Base64LineLength, "Base64 line length must be a multiple of four.");
        }

        WriteObject(new EmailWriterOptions(
            ConversionLossPolicy,
            UsePreservedRawSource.IsPresent,
            IncludeBccHeader.IsPresent,
            Base64LineLength,
            MaxNestedMessageDepth,
            MaxOutputBytes));
    }
}
