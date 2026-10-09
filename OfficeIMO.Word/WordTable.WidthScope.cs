using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        private sealed class LayoutSectionScope {
            internal LayoutSectionScope(WordSection section) { Section = section; }
            internal WordSection Section { get; }
        }

        /// <summary>
        /// Retains a known layout section when a header/footer part is shared by
        /// sections with different page geometry. The annotation is model state,
        /// not a serialized document attribute.
        /// </summary>
        internal void SetLayoutSection(WordSection section) {
            _table.RemoveAnnotations<LayoutSectionScope>();
            // Body ownership is determined from the current tree, including after a move.
            if (_table.Ancestors<Header>().Any() || _table.Ancestors<Footer>().Any())
                _table.AddAnnotation(new LayoutSectionScope(section));
        }

        // SDT/list wrappers do not change the table's containing cell.
        private TableCell? GetWidthHostCell() => _table.Ancestors<TableCell>().FirstOrDefault();

        private int EstimateContainingCellContentWidthInDxa() {
            TableCell? cell = GetWidthHostCell();
            if (cell != null) {
                int? width = EstimateCellContentWidthInDxa(_document, cell);
                if (width.HasValue) return width.Value;
            }
            return EstimateContentAreaWidthInDxa();
        }

        /// <summary>Resolves the section that supplies the table's page geometry.</summary>
        private WordSection? ResolveOwningSection() {
            var sections = _document.Sections;
            if (sections.Count == 0) return null;
            var main = _document._wordprocessingDocument.MainDocumentPart;
            var header = _table.Ancestors<Header>().FirstOrDefault()?.HeaderPart;
            var footer = _table.Ancestors<Footer>().FirstOrDefault()?.FooterPart;
            if (main != null && (header != null || footer != null)) {
                string id = header != null ? main.GetIdOfPart(header) : main.GetIdOfPart(footer!);
                var owners = WordStoryLayoutResolver.GetOwners(_document, id, header != null);
                var scope = _table.Annotation<LayoutSectionScope>();
                if (scope != null && owners.Contains(scope.Section)) return scope.Section;
                if (owners.Count == 0) return null;
                // A shared story has no unique page geometry after loading. Preserve its
                // existing definite grid rather than choosing one of several layout widths.
                int width = GetSectionTextWidth(owners[0]);
                return owners.All(section => GetSectionTextWidth(section) == width) ? owners[0] : null;
            }
            Body? body = main?.Document?.Body;
            if (body == null) return sections[0];
            OpenXmlElement owner = _table;
            while (owner.Parent != null && !ReferenceEquals(owner.Parent, body)) owner = owner.Parent;
            int index = 0;
            foreach (OpenXmlElement element in body.ChildElements) {
                if (ReferenceEquals(element, owner)) return sections[Math.Min(index, sections.Count - 1)];
                if (element is Paragraph paragraph && paragraph.ParagraphProperties?.SectionProperties != null && index < sections.Count - 1)
                    index++;
            }
            return sections[0];
        }

        private static int GetSectionTextWidth(WordSection section) =>
            Math.Max(0, (int)(section.PageSettings.Width ?? WordPageSizes.A4.WidthTwips)
                - (int)section.Margins.Left - (int)section.Margins.Right);

        private int? GetUnscopedStoryGridWidth() {
            if (!_table.Ancestors<Header>().Any() && !_table.Ancestors<Footer>().Any()) return null;
            if (ResolveOwningSection() != null) return null;
            long width = 0;
            foreach (GridColumn column in _table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
                         ?? Enumerable.Empty<GridColumn>()) {
                if (!int.TryParse(column.Width?.Value, out int value) || value <= 0) return null;
                width += value;
            }
            return width > 0 && width <= int.MaxValue ? (int)width : null;
        }
    }
}
