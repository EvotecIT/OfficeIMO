namespace OfficeIMO.Word {
    /// <summary>
    /// Base class for conversion option classes shared across OfficeIMO converters.
    /// </summary>
    public abstract class ConversionOptions : IConversionOptions {
        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string? FontFamily { get; set; }
    }
}
