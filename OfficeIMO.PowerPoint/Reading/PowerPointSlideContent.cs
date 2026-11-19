namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Content read from one slide by <see cref="PowerPointSlideReader"/>: its shapes in drawing order with their
    ///     text, tables and image references, and its speaker notes.
    /// </summary>
    public sealed class PowerPointSlideContent {
        internal PowerPointSlideContent(int number, bool hidden, IReadOnlyList<PowerPointShapeContent> shapes,
            string? notes) {
            Number = number;
            Hidden = hidden;
            Shapes = shapes;
            Notes = notes;
        }

        /// <summary>1-based slide number in presentation order.</summary>
        public int Number { get; }

        /// <summary>Whether the slide is hidden in slide shows.</summary>
        public bool Hidden { get; }

        /// <summary>
        ///     Shapes in drawing order. Group children follow their group, so the list is the slide's shape tree
        ///     flattened depth first.
        /// </summary>
        public IReadOnlyList<PowerPointShapeContent> Shapes { get; }

        /// <summary>Speaker notes text, or null when the slide has none or notes were not requested.</summary>
        public string? Notes { get; }

        /// <summary>Text of the first visible title or centered-title placeholder, or null when the slide has none.</summary>
        public string? Title {
            get {
                foreach (PowerPointShapeContent shape in Shapes) {
                    if (shape.IsVisible && shape.IsTitle) {
                        string text = shape.Text.Trim();
                        return text.Length == 0 ? null : text;
                    }
                }
                return null;
            }
        }

        /// <summary>Tables in drawing order.</summary>
        public IEnumerable<PowerPointTableContent> Tables {
            get {
                foreach (PowerPointShapeContent shape in Shapes) {
                    if (shape.Table != null) yield return shape.Table;
                }
            }
        }

        /// <summary>Picture and media image references in drawing order.</summary>
        public IEnumerable<PowerPointImageReference> Images {
            get {
                foreach (PowerPointShapeContent shape in Shapes) {
                    if (shape.Image != null) yield return shape.Image;
                }
            }
        }
    }

    /// <summary>A shape read by <see cref="PowerPointSlideReader"/>.</summary>
    public sealed class PowerPointShapeContent {
        internal PowerPointShapeContent(PowerPointShapeContentType kind, uint? id, string? name, string? altText,
            bool hidden, bool isVisible, int depth, PowerPointPlaceholderType? placeholderType,
            IReadOnlyList<PowerPointParagraphContent> paragraphs, PowerPointTableContent? table,
            PowerPointImageReference? image) {
            Kind = kind;
            Id = id;
            Name = name;
            AltText = altText;
            Hidden = hidden;
            IsVisible = isVisible;
            Depth = depth;
            PlaceholderType = placeholderType;
            Paragraphs = paragraphs;
            Table = table;
            Image = image;
        }

        /// <summary>Primary content carried by the shape.</summary>
        public PowerPointShapeContentType Kind { get; }

        /// <summary>Shape identifier from the non-visual drawing properties.</summary>
        public uint? Id { get; }

        /// <summary>Shape name.</summary>
        public string? Name { get; }

        /// <summary>Alternative text description.</summary>
        public string? AltText { get; }

        /// <summary>Whether the shape itself is marked hidden.</summary>
        public bool Hidden { get; }

        /// <summary>Whether the shape is shown: neither it nor any group containing it is hidden.</summary>
        public bool IsVisible { get; }

        /// <summary>Group nesting depth; 0 for shapes placed directly on the slide.</summary>
        public int Depth { get; }

        /// <summary>Placeholder type, when the shape is a typed placeholder.</summary>
        public PowerPointPlaceholderType? PlaceholderType { get; }

        /// <summary>Text body paragraphs of a text shape; empty for other shapes.</summary>
        public IReadOnlyList<PowerPointParagraphContent> Paragraphs { get; }

        /// <summary>Table content when <see cref="Kind"/> is <see cref="PowerPointShapeContentType.Table"/>.</summary>
        public PowerPointTableContent? Table { get; }

        /// <summary>Image reference for pictures and media posters.</summary>
        public PowerPointImageReference? Image { get; }

        /// <summary>Whether the shape is a title or centered-title placeholder.</summary>
        public bool IsTitle => PlaceholderType is PowerPointPlaceholderType.Title or PowerPointPlaceholderType.CenteredTitle;

        /// <summary>Paragraph texts joined with line breaks.</summary>
        public string Text => PowerPointParagraphContent.Join(Paragraphs);
    }

    /// <summary>A text paragraph read by <see cref="PowerPointSlideReader"/>.</summary>
    public sealed class PowerPointParagraphContent {
        internal PowerPointParagraphContent(IReadOnlyList<PowerPointInlineContent> inlines, int? level,
            string? bulletCharacter, bool isNumbered, PowerPointNumberingScheme? numberingScheme, int? numberingStartAt) {
            Inlines = inlines;
            Level = level;
            BulletCharacter = bulletCharacter;
            IsNumbered = isNumbered;
            NumberingScheme = numberingScheme;
            NumberingStartAt = numberingStartAt;
        }

        /// <summary>Runs, line breaks and fields in authored order.</summary>
        public IReadOnlyList<PowerPointInlineContent> Inlines { get; }

        /// <summary>Paragraph text; line breaks become <see cref="Environment.NewLine"/>.</summary>
        public string Text {
            get {
                if (Inlines.Count == 1) return Inlines[0].Text;
                return string.Concat(Inlines.Select(inline => inline.Text));
            }
        }

        /// <summary>List level (0-8) set directly on the paragraph.</summary>
        public int? Level { get; }

        /// <summary>Character bullet set directly on the paragraph.</summary>
        public string? BulletCharacter { get; }

        /// <summary>Whether the paragraph sets automatic numbering.</summary>
        public bool IsNumbered { get; }

        /// <summary>Automatic numbering scheme.</summary>
        public PowerPointNumberingScheme? NumberingScheme { get; }

        /// <summary>Explicit numbering start value.</summary>
        public int? NumberingStartAt { get; }

        internal static string Join(IReadOnlyList<PowerPointParagraphContent> paragraphs) {
            if (paragraphs.Count == 0) return string.Empty;
            if (paragraphs.Count == 1) return paragraphs[0].Text;
            return string.Join(Environment.NewLine, paragraphs.Select(paragraph => paragraph.Text));
        }
    }

    /// <summary>A run, line break or field inside a <see cref="PowerPointParagraphContent"/>.</summary>
    public sealed class PowerPointInlineContent {
        internal static readonly PowerPointInlineContent LineBreak =
            new(PowerPointParagraphInlineKind.LineBreak, Environment.NewLine, null, null);

        internal PowerPointInlineContent(PowerPointParagraphInlineKind kind, string text, Uri? hyperlink,
            string? fieldFallbackText) {
            Kind = kind;
            Text = text;
            Hyperlink = hyperlink;
            FieldFallbackText = fieldFallbackText;
        }

        /// <summary>Inline kind.</summary>
        public PowerPointParagraphInlineKind Kind { get; }

        /// <summary>Displayed text; a line break contributes <see cref="Environment.NewLine"/>.</summary>
        public string Text { get; }

        /// <summary>
        ///     Click hyperlink of a run. Internal slide links are returned as fragments such as <c>#slide-2</c>.
        /// </summary>
        public Uri? Hyperlink { get; }

        // A field without its own text element still has the text of its other children.
        internal string? FieldFallbackText { get; }
    }

    /// <summary>A table read by <see cref="PowerPointSlideReader"/>.</summary>
    public sealed class PowerPointTableContent {
        internal PowerPointTableContent(string? name, int columnCount, bool headerRow,
            IReadOnlyList<PowerPointTableRowContent> rows, bool hasGrid = true) {
            Name = name;
            ColumnCount = columnCount;
            HeaderRow = headerRow;
            Rows = rows;
            HasGrid = hasGrid;
        }

        // A table without a column grid, which the presentation model cannot size.
        internal bool HasGrid { get; }

        /// <summary>Name of the table's graphic frame.</summary>
        public string? Name { get; }

        /// <summary>Number of grid columns.</summary>
        public int ColumnCount { get; }

        /// <summary>Whether first-row (header) styling is enabled.</summary>
        public bool HeaderRow { get; }

        /// <summary>Rows in order.</summary>
        public IReadOnlyList<PowerPointTableRowContent> Rows { get; }
    }

    /// <summary>A table row read by <see cref="PowerPointSlideReader"/>.</summary>
    public sealed class PowerPointTableRowContent {
        internal PowerPointTableRowContent(IReadOnlyList<PowerPointTableCellContent> cells) {
            Cells = cells;
        }

        /// <summary>Cells in order, including cells covered by merges.</summary>
        public IReadOnlyList<PowerPointTableCellContent> Cells { get; }
    }

    /// <summary>A table cell read by <see cref="PowerPointSlideReader"/>.</summary>
    public sealed class PowerPointTableCellContent {
        internal PowerPointTableCellContent(IReadOnlyList<PowerPointParagraphContent> paragraphs) {
            Paragraphs = paragraphs;
        }

        /// <summary>Cell paragraphs.</summary>
        public IReadOnlyList<PowerPointParagraphContent> Paragraphs { get; }

        /// <summary>Paragraph texts joined with line breaks.</summary>
        public string Text => PowerPointParagraphContent.Join(Paragraphs);
    }

    /// <summary>The image a picture or media shape displays.</summary>
    public sealed class PowerPointImageReference {
        internal PowerPointImageReference(string relationshipId, Uri? partUri, string? contentType, Uri? externalUri) {
            RelationshipId = relationshipId;
            PartUri = partUri;
            ContentType = contentType;
            ExternalUri = externalUri;
        }

        /// <summary>Relationship identifier from the slide to the image.</summary>
        public string RelationshipId { get; }

        /// <summary>Package URI of an embedded image part, or null for a linked or missing image.</summary>
        public Uri? PartUri { get; }

        /// <summary>Content type of the embedded image part.</summary>
        public string? ContentType { get; }

        /// <summary>Target of a linked (external) image.</summary>
        public Uri? ExternalUri { get; }
    }
}
