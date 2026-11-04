using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>Retains native table/cell identity while layout supplies frame geometry and continuation.</summary>
    internal sealed class PowerPointLayoutNativeTable {
        private readonly PowerPointLayoutNativeTable _root;
        private readonly IReadOnlyList<A.TableRow> _rows;
        private readonly A.TableRow? _header;
        private readonly SlidePart? _sourcePart;
        private bool _used;

        internal PowerPointLayoutNativeTable(PowerPointTable table, bool headers) {
            Table = table;
            _root = this;
            _sourcePart = table.SlidePart;
            var rows = table.TableElement.Elements<A.TableRow>().ToArray();
            _header = headers ? rows[0] : null;
            _rows = headers ? rows.Skip(1).ToArray() : rows;
        }

        private PowerPointLayoutNativeTable(PowerPointLayoutNativeTable source, IReadOnlyList<A.TableRow> rows) {
            Table = source.Table; _root = source._root; _header = source._header; _rows = rows; _sourcePart = source._sourcePart;
        }

        internal PowerPointTable Table { get; }
        internal IReadOnlyList<string> Headers => _header == null
            ? Enumerable.Range(1, Table.Columns).Select(i => "Column " + i).ToArray() : Values(_header);
        internal IReadOnlyList<IReadOnlyList<string>> Rows => _rows.Select(Values).ToArray();
        private IReadOnlyList<string> Values(A.TableRow row) => row.Elements<A.TableCell>().Select(c => Table.WrapCell(c).Text).ToArray();

        internal PowerPointLayoutNativeTable Slice(IReadOnlyList<IReadOnlyList<string>> sourceRows, IReadOnlyList<IReadOnlyList<string>> selected) {
            var indices = selected.Select(row => Enumerable.Range(0, sourceRows.Count).Single(i => ReferenceEquals(sourceRows[i], row)));
            return new PowerPointLayoutNativeTable(this, indices.Select(i => _rows[i]).ToArray());
        }

        internal PowerPointTable Bind(PowerPointSlide slide, PowerPointTable generated) {
            var root = _root;
            var result = root._used ? generated : Table;
            var template = generated.TableElement;
            var defaults = template.Elements<A.TableRow>().ToArray();
            var rows = new List<A.TableRow>();
            var sources = new List<SlidePart?>();
            if (_header != null) {
                sources.Add(_header.Ancestors<P.Slide>().FirstOrDefault()?.SlidePart ?? _sourcePart);
                rows.Add(root._used ? (A.TableRow)_header.CloneNode(true) : _header);
            }
            rows.AddRange(_rows);
            sources.AddRange(_rows.Select(row => row.Ancestors<P.Slide>().FirstOrDefault()?.SlidePart ?? _sourcePart));
            for (int i = 0; i < rows.Count; i++) {
                var row = rows[i];
                var source = sources[i];
                if (!ReferenceEquals(source, slide.SlidePart)) PowerPointLayoutNativeText.RehomeRelationships(row, source, slide.SlidePart);
                row.Height = defaults[i].Height;
                var cells = row.Elements<A.TableCell>().ToArray();
                var fallback = defaults[i].Elements<A.TableCell>().ToArray();
                for (int c = 0; c < cells.Length; c++) {
                    PowerPointLayoutNativeText.MergeMissing(cells[c].TableCellProperties ??= new A.TableCellProperties(), fallback[c].TableCellProperties);
                    var paragraphs = cells[c].TextBody!.Elements<A.Paragraph>().ToArray();
                    var paragraphDefault = fallback[c].TextBody!.Elements<A.Paragraph>().First();
                    foreach (var paragraph in paragraphs) {
                        PowerPointLayoutNativeText.MergeMissing(paragraph.ParagraphProperties ??= new A.ParagraphProperties(), paragraphDefault.ParagraphProperties);
                        var runDefault = paragraphDefault.GetFirstChild<A.Run>()?.RunProperties;
                        foreach (var run in paragraph.Elements<A.Run>())
                            PowerPointLayoutNativeText.MergeMissing(run.RunProperties ??= new A.RunProperties(), runDefault);
                    }
                }
            }

            var properties = (A.TableProperties)Table.TableElement.TableProperties!.CloneNode(true);
            var grid = (A.TableGrid)template.TableGrid!.CloneNode(true);
            var transform = (P.Transform)((P.GraphicFrame)generated.Element).Transform!.CloneNode(true);
            foreach (var row in rows) if (row.Parent != null) row.Remove();
            var target = result.TableElement;
            target.RemoveAllChildren<A.TableRow>();
            target.TableProperties = properties;
            target.TableGrid = grid;
            foreach (var row in rows) target.Append(row);
            ((P.GraphicFrame)result.Element).Transform = transform;
            if (!ReferenceEquals(result, generated)) {
                slide.RemoveShape(generated);
                slide.AdoptLayoutShape(result);
            }
            root._used = true;
            return result;
        }
    }
}
