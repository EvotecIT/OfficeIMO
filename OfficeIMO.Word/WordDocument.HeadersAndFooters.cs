namespace OfficeIMO.Word {
    // Exposes the first section's page setup from the document root.
    public partial class WordDocument {
        /// <summary>
        /// Gets the page size, margins, orientation and running headers and footers of the first section.
        /// </summary>
        /// <remarks>
        /// When a document contains multiple sections, use
        /// <c>Sections[wantedSection].PageSetup</c> to target a specific section.
        /// </remarks>
        public WordPageSetup PageSetup => this.Sections[0].PageSetup;
    }
}
