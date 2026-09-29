namespace OfficeIMO.Word {
    /// <summary>
    /// Controls deterministic behavior used while refreshing supported Word fields.
    /// </summary>
    public sealed class WordFieldUpdateOptions {
        /// <summary>
        /// Gets an options instance that uses the current local date and time for DATE and TIME fields.
        /// </summary>
        public static WordFieldUpdateOptions Default => new WordFieldUpdateOptions();

        /// <summary>
        /// Gets or sets the date and time used for DATE and TIME field refresh.
        /// When unset, OfficeIMO uses the current local date and time.
        /// </summary>
        public DateTime? CurrentDateTime { get; set; }

        /// <summary>
        /// When true, refreshes only DATE and TIME fields and leaves all other field results unchanged.
        /// </summary>
        public bool DateTimeFieldsOnly { get; set; }
    }
}
