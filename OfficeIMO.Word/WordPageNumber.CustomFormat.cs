namespace OfficeIMO.Word {
    public partial class WordPageNumber {
        /// <summary>
        /// Gets or sets the custom format of the page number field.
        /// </summary>
        public string CustomFormat {
            get { return Field!.GetCustomFormat(); }
            set { Field!.SetCustomFormat(value); }
        }
    }
}
