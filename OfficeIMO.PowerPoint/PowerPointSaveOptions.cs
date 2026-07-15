namespace OfficeIMO.PowerPoint {
    /// <summary>Controls whether conversion to a less expressive PowerPoint format can omit content.</summary>
    public enum PowerPointConversionLossPolicy {
        /// <summary>Reject a conversion when known content or formatting cannot be represented.</summary>
        Block,

        /// <summary>Allow a conversion after known losses have been reported by preflight.</summary>
        Allow
    }

    /// <summary>Controls PowerPoint save and conversion behavior.</summary>
    public sealed class PowerPointSaveOptions {
        /// <summary>Gets or sets how known conversion loss is handled.</summary>
        public PowerPointConversionLossPolicy LossPolicy { get; set; } = PowerPointConversionLossPolicy.Block;
    }
}
