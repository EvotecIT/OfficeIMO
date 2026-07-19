namespace OfficeIMO.Word {
    /// <summary>
    /// Identifies the portion of a source Word body block rendered on one visual page.
    /// </summary>
    public sealed class WordDocumentVisualFragment {
        internal WordDocumentVisualFragment(
            int sectionIndex,
            int blockIndex,
            string kind,
            string text,
            WordDocumentVisualRegion? region) {
            SectionIndex = sectionIndex;
            BlockIndex = blockIndex;
            Kind = kind;
            Text = text;
            Region = region;
        }

        /// <summary>Zero-based source section index.</summary>
        public int SectionIndex { get; }

        /// <summary>Zero-based body-block index within the source section.</summary>
        public int BlockIndex { get; }

        /// <summary>Source block kind, such as paragraph or table.</summary>
        public string Kind { get; }

        /// <summary>Best-effort visible text painted for this source block on the page.</summary>
        public string Text { get; }

        /// <summary>Best-effort page region occupied by the rendered source block.</summary>
        public WordDocumentVisualRegion? Region { get; }
    }

    /// <summary>
    /// Rectangular region in the top-left, point-based coordinate space of a Word visual page.
    /// </summary>
    public sealed class WordDocumentVisualRegion {
        internal WordDocumentVisualRegion(double x, double y, double width, double height) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        /// <summary>Left coordinate in points.</summary>
        public double X { get; }

        /// <summary>Top coordinate in points.</summary>
        public double Y { get; }

        /// <summary>Region width in points.</summary>
        public double Width { get; }

        /// <summary>Region height in points.</summary>
        public double Height { get; }
    }
}
