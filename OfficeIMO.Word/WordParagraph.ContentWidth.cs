using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Estimates the authored content width available to an inline drawing in this paragraph, in points.
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
        public double GetContentWidthPoints() {
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

            double before = Math.Max(0D, IndentationBeforePoints ?? 0D);
            double after = Math.Max(0D, IndentationAfterPoints ?? 0D);
            double firstLine = IndentationHangingPoints > 0D ? 0D : Math.Max(0D, IndentationFirstLinePoints ?? 0D);
            double width = containerWidth - before - after - firstLine;
            if (width <= 0D) {
                throw new InvalidOperationException("The paragraph margins and indents leave no positive content width. Use explicit drawing dimensions.");
            }
            return width;
        }

        private double GetSectionPageContentWidthPoints(WordSection section) {
            double gutter = _document.Settings.GutterAtTop ? 0D : section.Margins.Gutter;
            return ((section.PageSettings.Width ?? WordPageSizes.A4.WidthTwips)
                - (double)section.Margins.Left - section.Margins.Right - gutter) / 20D;
        }

        private double GetSectionColumnContentWidthPoints(WordSection section) {
            double width = GetSectionPageContentWidthPoints(section);
            IReadOnlyList<WordSectionColumn> definitions = section.ColumnDefinitions;
            if (definitions.Count > 0) {
                return Math.Min(width, definitions.Min(column => column.WidthTwips) / 20D);
            }
            int count = Math.Max(1, section.ColumnCount ?? 1);
            double gap = Math.Max(0, section.ColumnsSpace ?? 720) / 20D;
            return (width - gap * (count - 1)) / count;
        }
    }
}
