namespace OfficeIMO.Word {
    public partial class WordSmartArt {
        /// <summary>Reads persisted inline or anchored package geometry for this SmartArt diagram.</summary>
        /// <param name="snapshot">The package-geometry snapshot when the diagram has a supported frame.</param>
        /// <returns><see langword="true"/> when persisted layout evidence was available.</returns>
        public bool TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot snapshot) =>
            WordDrawingLayoutReader.TryRead(_drawing, out snapshot);
    }
}
