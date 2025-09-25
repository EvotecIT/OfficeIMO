namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Layout primitives for SheetComposer (titles, sections, lists, KPIs, score, references).
    /// </summary>
    public sealed partial class SheetComposer {
        /// <summary>Adds vertical spacing by advancing the current row.</summary>
        /// <param name="rows">Number of rows to skip; when negative uses the theme default.</param>
        public SheetComposer Spacer(int rows = -1) { _row += rows > 0 ? rows : _theme.DefaultSpacingRows; return this; }

        /// <summary>
        /// Advances the composer by an exact number of rows without applying themed spacing semantics.
        /// Useful after manual cell writes where the composer did not track row changes.
        /// </summary>
        public SheetComposer AdvanceRows(int rows) {
            if (rows <= 0) return this;
            _row += rows;
            return this;
        }

        /// <summary>Writes a title (bold with themed background) and optional subtitle.</summary>
        public SheetComposer Title(string text, string? subtitle = null) {
            if (string.IsNullOrWhiteSpace(text)) return this;
            Sheet.Cell(_row, 1, text);
            Sheet.CellBold(_row, 1, true);
            Sheet.CellBackground(_row, 1, _theme.SectionHeaderFillHex);
            _row++;
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                Sheet.Cell(_row, 1, subtitle);
                _row++;
            }
            return Spacer();
        }

        /// <summary>Writes a section header (bold with themed background).</summary>
        public SheetComposer Section(string text) {
            Sheet.Cell(_row, 1, text);
            Sheet.CellBold(_row, 1, true);
            Sheet.CellBackground(_row, 1, _theme.SectionHeaderFillHex);
            _row++;
            return this;
        }

        /// <summary>Writes a paragraph-like line of text.</summary>
        public SheetComposer Paragraph(string text, int widthColumns = 6) {
            if (string.IsNullOrEmpty(text)) return this;
            Sheet.Cell(_row, 1, text);
            _row++;
            return this;
        }

        /// <summary>
        /// Inserts a simple callout (admonition) band consisting of a bold title row and a body row.
        /// </summary>
        public SheetComposer Callout(string kind, string title, string body, int widthColumns = 8) {
            string fill = kind?.Trim().ToLowerInvariant() switch {
                "success" => "#D4EDDA",
                "warning" => "#FFF3CD",
                "error" => "#F8D7DA",
                "critical" => "#F8D7DA",
                _ => "#E8F4FF"
            };

            if (!string.IsNullOrWhiteSpace(title)) {
                Sheet.Cell(_row, 1, title);
                Sheet.CellBold(_row, 1, true);
                for (int c = 1; c <= Math.Max(1, widthColumns); c++) Sheet.CellBackground(_row, c, fill);
                _row++;
            }

            if (!string.IsNullOrWhiteSpace(body)) {
                string text = body;
                if (!text.Contains("\n") && text.Length > 120) {
                    int cut = 120; text = text.Insert(cut, "\n");
                }
                Sheet.Cell(_row, 1, text);
                for (int c = 1; c <= Math.Max(1, widthColumns); c++) Sheet.CellBackground(_row, c, fill);
                _row++;
            }
            return Spacer();
        }

        /// <summary>Alias for <see cref="PropertiesGrid"/>.</summary>
        public SheetComposer DefinitionList(IEnumerable<(string Key, object? Value)> items, int columns = 2)
            => PropertiesGrid(items, columns);

        /// <summary>Renders a compact grid of key/value pairs.</summary>
        public SheetComposer PropertiesGrid(IEnumerable<(string Key, object? Value)> properties, int columns = 2) {
            if (properties == null) return this;
            var list = new List<(string Key, object? Value)>(properties);
            if (list.Count == 0) return this;
            int idx = 0;
            while (idx < list.Count) {
                int col = 1;
                for (int c = 0; c < columns && idx < list.Count; c++, idx++) {
                    var (k, v) = list[idx];
                    Sheet.Cell(_row, col, k);
                    Sheet.CellBold(_row, col, true);
                    Sheet.CellBackground(_row, col, _theme.KeyFillHex);
                    Sheet.Cell(_row, col + 1, v ?? string.Empty);
                    col += 2;
                }
                _row++;
            }
            return Spacer();
        }

        /// <summary>Writes a simple bulleted list, one item per row.</summary>
        public SheetComposer BulletedList(IEnumerable<string> items) {
            if (items == null) return this;
            foreach (var item in items) {
                Sheet.Cell(_row, 1, $"• {item}");
                _row++;
            }
            return Spacer();
        }

        /// <summary>Bulleted list with background fill per item row.</summary>
        public SheetComposer BulletedListWithFill(IEnumerable<string> items, string fillHex) {
            if (items == null) return this;
            int start = _row;
            BulletedList(items);
            int end = _row - 1 - _theme.DefaultSpacingRows;
            if (end >= start)
                for (int r = start; r <= end; r++)
                    Sheet.CellBackground(r, 1, fillHex);
            return this;
        }

        /// <summary>Renders a compact KPI row of label/value pairs.</summary>
        public SheetComposer KpiRow(IEnumerable<(string Label, object? Value)> kpis, int perRow = 3, string? labelFillHex = null) {
            if (kpis == null) return this;
            var list = new List<(string Label, object? Value)>(kpis);
            if (list.Count == 0) return this;

            int idx = 0; string fill = labelFillHex ?? _theme.KeyFillHex;
            while (idx < list.Count) {
                int col = 1; int rendered = 0;
                for (; rendered < perRow && idx + rendered < list.Count; rendered++) {
                    var (label, _) = list[idx + rendered];
                    Sheet.Cell(_row, col, label);
                    Sheet.CellBold(_row, col, true);
                    Sheet.CellBackground(_row, col, fill);
                    Sheet.CellAlign(_row, col, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    col++;
                }
                _row++;

                col = 1;
                for (int i = 0; i < rendered; i++) {
                    var (_, val) = list[idx + i];
                    Sheet.Cell(_row, col, val ?? string.Empty);
                    Sheet.CellBold(_row, col, true);
                    Sheet.CellAlign(_row, col, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    col++;
                }
                _row++;
                idx += rendered;
            }
            return Spacer();
        }

        /// <summary>Writes a labeled numeric score with a data bar visualization.</summary>
        public SheetComposer Score(string label, double value, double min = 0, double max = 10) {
            Sheet.Cell(_row, 1, label);
            Sheet.CellBold(_row, 1, true);
            Sheet.Cell(_row, 2, value);
            string range = $"B{_row}:B{_row}";
            Sheet.AddConditionalDataBar(range, SixLabors.ImageSharp.Color.LightGreen);
            _row++;
            return Spacer();
        }

        /// <summary>Writes a simple References section with each URL as a hyperlink.</summary>
        public SheetComposer References(IEnumerable<string> urls) {
            var list = urls is null ? null : new List<string>(urls);
            if (list != null && list.Count > 0) {
                Section("References");
                foreach (var url in list) { Sheet.SetHyperlinkSmart(_row, 1, url); _row++; }
                Spacer();
            }
            return this;
        }
    }
}
