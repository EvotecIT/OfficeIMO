using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static ExcelWriterOptions CreateRowWriteOptions(ExcelWriterOptions? options) {
            if (options == null) {
                return new ExcelWriterOptions { UseSharedStrings = false };
            }

            return new ExcelWriterOptions {
                SheetName = options.SheetName,
                IncludeHeaders = options.IncludeHeaders,
                CreateTable = options.CreateTable,
                TableName = options.TableName,
                TableStyle = options.TableStyle,
                IncludeAutoFilter = options.IncludeAutoFilter,
                AutoFit = options.AutoFit,
                UseCellValueNumberFormats = options.UseCellValueNumberFormats,
                IncludeCellReferences = options.IncludeCellReferences,
                UseSharedStrings = options.UseSharedStrings,
                RequireStreaming = options.RequireStreaming,
                DateSystem = options.DateSystem,
                Styles = options.Styles,
                DefaultRowStyle = options.DefaultRowStyle,
                ColumnStyles = options.ColumnStyles
            };
        }

        private static void ValidateTabularHeaders(IReadOnlyList<string> headers) {
            var usedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++) {
                string header = headers[columnIndex] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(header)) {
                    throw new ArgumentException("Column headers must not be empty.", nameof(headers));
                }
                if (!usedHeaders.Add(header)) {
                    throw new ArgumentException("Column headers must be unique.", nameof(headers));
                }
            }
        }

    }
}
