using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Estimates the authored content width available to an inline drawing in this paragraph as a physical length.
        /// </summary>
        /// <returns>A positive width after the containing cell or section geometry and direct paragraph indents.</returns>
        /// <remarks>
        /// Uses the containing table cell's existing width estimate, or the owning section's page margins and
        /// column settings. Unequal flowing columns use their smallest authored width because Word's final
        /// pagination determines which column contains the paragraph. Header/footer stories use the full page
        /// content width and must have a unique width across their owning sections. This is an authored geometry
        /// estimate, not Word's measured automatic table layout, pagination, or remaining space on a text line.
        /// Explicit drawing dimensions can be used where the authored geometry is insufficient.
        /// </remarks>
        /// <exception cref="InvalidOperationException">
        /// The paragraph is detached, belongs to an unsupported text box, has ambiguous story geometry,
        /// or its margins and indents leave no positive content width.
        /// </exception>
        public OfficeLength GetContentWidth() {
            if (_paragraph.Parent == null) {
                throw new InvalidOperationException("The paragraph must be attached to a document before its content width can be estimated.");
            }
            if (ResolveParent(_document, _paragraph) is WordTextBox) {
                throw new InvalidOperationException("Text-box content width requires explicit drawing dimensions.");
            }

            Header? header = _paragraph.Ancestors<Header>().FirstOrDefault();
            Footer? footer = _paragraph.Ancestors<Footer>().FirstOrDefault();
            double containerWidth;
            if (header != null || footer != null) {
                var main = _document._wordprocessingDocument.MainDocumentPart;
                var part = header != null
                    ? (DocumentFormat.OpenXml.Packaging.OpenXmlPart?)header.HeaderPart
                    : footer!.FooterPart;
                if (main == null || part == null) {
                    throw new InvalidOperationException("The header/footer paragraph has no owning section geometry. Use explicit drawing dimensions.");
                }
                var owners = WordStoryLayoutResolver.GetOwners(_document, main.GetIdOfPart(part), header != null);
                if (owners.Count == 0) {
                    throw new InvalidOperationException("The header/footer paragraph has no owning section geometry. Use explicit drawing dimensions.");
                }
                containerWidth = GetSectionPageContentWidthPoints(owners[0]);
                if (owners.Any(section => GetSectionPageContentWidthPoints(section) != containerWidth)) {
                    throw new InvalidOperationException("The shared header/footer belongs to sections with different content widths. Use explicit drawing dimensions.");
                }
            } else {
                WordSection? section = FindBodySection(_document, _paragraph);
                if (section == null) {
                    throw new InvalidOperationException("The paragraph has no supported owning section geometry. Use explicit drawing dimensions.");
                }
                containerWidth = GetSectionColumnContentWidthPoints(section);
            }

            TableCell? cell = _paragraph.Ancestors<TableCell>().FirstOrDefault();
            if (cell != null) {
                int? cellWidth = WordTable.EstimateCellContentWidthInDxa(_document, cell);
                if (!cellWidth.HasValue || cellWidth.Value <= 0) {
                    throw new InvalidOperationException("The containing cell has no usable width estimate. Use explicit drawing dimensions.");
                }
                containerWidth = Math.Min(containerWidth, cellWidth.Value / 20D);
            }

            double before = Math.Max(0D, IndentationBefore?.ToPoints() ?? 0D);
            double after = Math.Max(0D, IndentationAfter?.ToPoints() ?? 0D);
            double firstLine = IndentationHanging?.ToPoints() > 0D ? 0D : Math.Max(0D, IndentationFirstLine?.ToPoints() ?? 0D);
            double width = containerWidth - before - after - firstLine;
            if (width <= 0D) {
                throw new InvalidOperationException("The paragraph margins and indents leave no positive content width. Use explicit drawing dimensions.");
            }
            return OfficeLength.Points(width);
        }

        private double GetSectionPageContentWidthPoints(WordSection section) {
            double gutter = _document.Settings.GutterAtTop ? 0D : checked((uint)(section.Margins.Gutter.GetValueOrDefault().ToTwips()));
            return ((checked((uint?)(section.PageSettings.Width?.ToTwips())) ?? checked((uint)(WordPageSizes.A4.Width.ToTwips())))
                - (double)checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips())) - checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips())) - gutter) / 20D;
        }

        private double GetSectionColumnContentWidthPoints(WordSection section) {
            double width = GetSectionPageContentWidthPoints(section);
            IReadOnlyList<WordSectionColumn> definitions = section.ColumnDefinitions;
            if (definitions.Count > 0) {
                return Math.Min(width, definitions.Min(column => column.Width.ToPoints()));
            }
            int count = Math.Max(1, section.ColumnCount ?? 1);
            double gap = Math.Max(0D, section.ColumnsSpace?.ToPoints() ?? 36D);
            return (width - gap * (count - 1)) / count;
        }
    }
}
