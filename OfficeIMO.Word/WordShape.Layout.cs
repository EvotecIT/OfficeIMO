namespace OfficeIMO.Word {
    public partial class WordShape {
        /// <summary>Reads persisted inline or anchored package geometry for this DrawingML shape.</summary>
        /// <param name="snapshot">The package-geometry snapshot when this is a DrawingML shape with a supported frame.</param>
        /// <returns><see langword="true"/> for DrawingML shapes with persisted frame geometry; VML-only shapes return false.</returns>
        public bool TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot snapshot) =>
            WordDrawingLayoutReader.TryRead(_drawing, out snapshot);
    }
}
