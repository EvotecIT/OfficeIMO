using System;

namespace OfficeIMO;

/// <summary>Shared resource limits for one local document workflow.</summary>
public sealed class OfficeOperationLimits {
    /// <summary>Maximum primary or comparison input size.</summary>
    public long MaximumInputBytes { get; set; } = 256L * 1024L * 1024L;

    /// <summary>Maximum generated artifact size.</summary>
    public long MaximumOutputBytes { get; set; } = 512L * 1024L * 1024L;

    /// <summary>Maximum characters in each XML part read by built-in Open XML, Draw and Visio conversions, or in a complete legacy Visio XML source. Default: 10 MiB.</summary>
    public long MaximumXmlCharactersInPart { get; set; } = 10L * 1024L * 1024L;

    internal OfficeOperationLimits CloneAndValidate() {
        if (MaximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumInputBytes));
        if (MaximumOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumOutputBytes));
        if (MaximumXmlCharactersInPart <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumXmlCharactersInPart));
        return new OfficeOperationLimits {
            MaximumInputBytes = MaximumInputBytes,
            MaximumOutputBytes = MaximumOutputBytes,
            MaximumXmlCharactersInPart = MaximumXmlCharactersInPart
        };
    }
}
