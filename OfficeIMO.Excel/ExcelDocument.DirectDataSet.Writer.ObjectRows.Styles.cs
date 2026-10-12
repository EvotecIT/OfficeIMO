namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        public sealed partial class ExcelTabularRowWriter {
            private ExcelTabularStylePlan? _declaredStyles;
            private ExcelTabularStylePlan.DeclaredStyle? _rowStyle;
            private ExcelTabularStylePlan.DeclaredStyle? _nextCellStyle;
            private bool _rowOpened;
            private bool _rowActive;

            /// <summary>
            /// Selects a declared default for the current row before its first cell. Null clears
            /// the configured row default. The selected complete style overrides column defaults.
            /// </summary>
            public ExcelTabularRowWriter SetRowStyle(string? name) {
                EnsureActiveRow();
                if (_rowOpened) throw new InvalidOperationException("SetRowStyle must be called before writing the first cell in the row.");
                _rowStyle = ResolveDeclaredStyle(name);
                return this;
            }

            /// <summary>
            /// Selects a declared complete style for the next cell only. Null restores normal
            /// row/column inheritance. An empty declared definition explicitly applies Normal formatting.
            /// </summary>
            public ExcelTabularRowWriter SetNextCellStyle(string? name) {
                EnsureActiveRow();
                if (_columnIndex >= _cellReferencePrefixes.Length) throw new InvalidOperationException("The row has no remaining cells to style.");
                _nextCellStyle = ResolveDeclaredStyle(name);
                return this;
            }

            /// <summary>Writes a genuine blank cell, with the selected style or inherited row/column default.</summary>
            public ExcelTabularRowWriter WriteBlank() {
                BeginCell();
                _writer.Write("/>");
                return this;
            }

            private ExcelTabularStylePlan.DeclaredStyle? ResolveDeclaredStyle(string? name) {
                if (name == null) return null;
                if (_declaredStyles == null) throw new KeyNotFoundException("Declare named styles in ExcelTabularWriteOptions.Styles before selecting a style.");
                return _declaredStyles.Resolve(name);
            }

            private void EnsureActiveRow() {
                if (!_rowActive) throw new InvalidOperationException("The row writer can only be used while its row callback is running.");
            }

            private void OpenRow() {
                if (_rowOpened) return;
                if (_includeCellReferences) {
                    if (_rowStyle == null) {
                        WriteReferenceStart("<row r=\"", "\">");
                        _rowOpened = true;
                        return;
                    }
                    WriteReferenceStart("<row r=\"", "\"");
                } else {
                    _writer.Write("<row");
                }
                if (_rowStyle != null) {
                    _writer.Write(_rowStyle.Attribute);
                    _writer.Write(" customFormat=\"1\"");
                }
                _writer.Write('>');
                _rowOpened = true;
            }

            private string? GetDeclaredCellAttribute(int column, string? temporalAttribute) {
                ExcelTabularStylePlan.DeclaredStyle? explicitStyle = _nextCellStyle;
                _nextCellStyle = null;
                ExcelTabularStylePlan.DeclaredStyle? effective = explicitStyle ?? _rowStyle ?? _declaredStyles?.Columns[column];
                if (effective == null) return temporalAttribute;
                // Excel applies row/column defaults to cells absent from sheetData. Authored cells
                // also need the resolved XF; automatic temporal formats preserve that visual style.
                return effective.ForValue(temporalAttribute);
            }
        }
    }
}
