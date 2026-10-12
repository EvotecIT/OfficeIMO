namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            internal sealed partial class DirectSharedStringTable {
                private readonly List<string>? _streamingValues;

                private DirectSharedStringTable() {
                    _indexes = new Dictionary<string, int>(StringComparer.Ordinal);
                    _streamingValues = new List<string>();
                    Values = _streamingValues;
                }

                /// <summary>
                /// Interns text while rows are written once. Only unique values are retained;
                /// the completed table must be serialized after all worksheet cells.
                /// </summary>
                internal static DirectSharedStringTable CreateStreaming() => new DirectSharedStringTable();

                private bool TryCollectIndex(string value, out int index) {
                    CoerceValueHelper.ValidateSharedStringLength(value, "value");
                    if (!_indexes.TryGetValue(value, out index)) {
                        index = _streamingValues!.Count;
                        _indexes.Add(value, index);
                        _streamingValues.Add(value);
                    }
                    TotalStringReferences = checked(TotalStringReferences + 1);
                    return true;
                }
            }
        }
    }
}
