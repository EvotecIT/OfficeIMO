using System.Globalization;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private static void WriteTable(ZipArchive archive, DirectDataSetSheetModel sheet) {
                int columnCount = sheet.Table.ColumnCount;
                var builder = new StringBuilder(512 + (columnCount * 72));
                string sheetIndexText = InvariantNumberText.Get(sheet.Index);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"");
                builder.Append(sheetIndexText);
                builder.Append("\" name=\"");
                AppendEscaped(builder, sheet.TableName!);
                builder.Append("\" displayName=\"");
                AppendEscaped(builder, sheet.TableName!);
                builder.Append("\" ref=\"");
                AppendEscaped(builder, sheet.Range);
                builder.Append("\" headerRowCount=\"");
                builder.Append(sheet.IncludeHeaders ? "1" : "0");
                builder.Append("\" totalsRowShown=\"0\">");
                if (sheet.IncludeAutoFilter && sheet.IncludeHeaders) {
                    builder.Append("<autoFilter ref=\"");
                    AppendEscaped(builder, sheet.Range);
                    builder.Append("\"/>");
                }

                builder.Append("<tableColumns count=\"");
                builder.Append(InvariantNumberText.Get(columnCount));
                builder.Append("\">");
                for (int i = 0; i < columnCount; i++) {
                    string columnIndexText = InvariantNumberText.Get(i + 1);
                    builder.Append("<tableColumn id=\"");
                    builder.Append(columnIndexText);
                    builder.Append("\" name=\"");
                    string columnName = sheet.Table.GetColumnName(i);
                    if (string.IsNullOrWhiteSpace(columnName)) {
                        builder.Append("Column");
                        builder.Append(columnIndexText);
                    } else {
                        AppendEscaped(builder, columnName);
                    }

                    builder.Append("\"/>");
                }

                builder.Append("</tableColumns><tableStyleInfo name=\"");
                builder.Append(sheet.TableStyle.ToString());
                builder.Append("\" showFirstColumn=\"");
                builder.Append(sheet.ShowFirstColumn ? "1" : "0");
                builder.Append("\" showLastColumn=\"");
                builder.Append(sheet.ShowLastColumn ? "1" : "0");
                builder.Append("\" showRowStripes=\"");
                builder.Append(sheet.ShowRowStripes ? "1" : "0");
                builder.Append("\" showColumnStripes=\"");
                builder.Append(sheet.ShowColumnStripes ? "1" : "0");
                builder.Append("\"/></table>");
                WriteTextEntry(archive, "xl/tables/table" + sheetIndexText + ".xml", builder.ToString());
            }

            internal enum DirectCellValueKind {
                Object,
                Formula,
                String,
                Boolean,
                DateTime,
                DateTimeOffset,
                TimeSpan,
                Double,
                Float,
                Decimal,
                SByte,
                Byte,
                Int16,
                UInt16,
                Int32,
                UInt32,
                Int64,
                UInt64,
#if NET6_0_OR_GREATER
                DateOnly,
                TimeOnly
#endif
            }

            internal sealed class DirectStylePlan {
                private const string DateTimeFormatCode = "yyyy-mm-dd hh:mm";
                private const string TimeFormatCode = "[h]:mm:ss";
                private readonly Dictionary<string, string> _styleAttributeByFormat;

                private DirectStylePlan(List<string> customNumberFormats, Dictionary<string, string> styleAttributeByFormat) {
                    CustomNumberFormats = customNumberFormats;
                    _styleAttributeByFormat = styleAttributeByFormat;
                }

                internal IReadOnlyList<string> CustomNumberFormats { get; }

                internal static DirectStylePlan Create(DirectDataSetWorkbookModel model) {
                    var customNumberFormats = new List<string>();
                    var styleAttributeByFormat = new Dictionary<string, string>(StringComparer.Ordinal);
                    styleAttributeByFormat[DateTimeFormatCode] = DateStyleAttribute;
                    styleAttributeByFormat[TimeFormatCode] = TimeStyleAttribute;
                    for (int sheetIndex = 0; sheetIndex < model.Sheets.Count; sheetIndex++) {
                        var formats = model.Sheets[sheetIndex].ColumnNumberFormats;
                        if (formats != null) {
                            for (int i = 0; i < formats.Count; i++) {
                                AddCustomNumberFormat(formats[i], customNumberFormats, styleAttributeByFormat);
                            }
                        }

                        var overlayCells = model.Sheets[sheetIndex].Metadata?.OverlayCells;
                        if (overlayCells != null) {
                            for (int i = 0; i < overlayCells.Count; i++) {
                                if (overlayCells[i].IsDeleted) {
                                    continue;
                                }

                                AddCustomNumberFormat(overlayCells[i].NumberFormat, customNumberFormats, styleAttributeByFormat);
                            }
                        }
                    }

                    return new DirectStylePlan(customNumberFormats, styleAttributeByFormat);
                }

                private static void AddCustomNumberFormat(string? format, List<string> customNumberFormats, Dictionary<string, string> styleAttributeByFormat) {
                    if (string.IsNullOrWhiteSpace(format) || styleAttributeByFormat.ContainsKey(format!)) {
                        return;
                    }

                    string styleAttribute = " s=\"" + InvariantNumberText.Get(5 + customNumberFormats.Count) + "\"";
                    styleAttributeByFormat.Add(format!, styleAttribute);
                    customNumberFormats.Add(format!);
                }

                internal string? GetStyleAttribute(string numberFormat)
                    => _styleAttributeByFormat.TryGetValue(numberFormat, out string? styleAttribute)
                        ? styleAttribute
                        : null;
            }

            internal sealed class ExtendedDirectWritePlan {
                internal ExtendedDirectWritePlan(
                    DirectDataSetWorkbookModel model,
                    DirectStylePlan stylePlan,
                    DirectColumnWritePlan[] columnWritePlans,
                    DirectSharedStringTable? sharedStrings) {
                    Model = model;
                    StylePlan = stylePlan;
                    ColumnWritePlans = columnWritePlans;
                    SharedStrings = sharedStrings;
                }

                internal DirectDataSetWorkbookModel Model { get; }

                internal DirectStylePlan StylePlan { get; }

                internal DirectColumnWritePlan[] ColumnWritePlans { get; }

                internal DirectSharedStringTable? SharedStrings { get; }

                internal bool HasSharedStrings => SharedStrings != null;
            }

            internal readonly struct DirectColumnWritePlan {
                internal DirectColumnWritePlan(DirectCellValueKind[] cellValueKinds, string?[]? styleAttributes, bool[]? valueStyleColumns) {
                    CellValueKinds = cellValueKinds;
                    StyleAttributes = styleAttributes;
                    ValueStyleColumns = valueStyleColumns;
                }

                internal DirectCellValueKind[] CellValueKinds { get; }

                internal string?[]? StyleAttributes { get; }

                internal bool[]? ValueStyleColumns { get; }
            }

            internal sealed class DirectSharedStringTable {
                private const int MinimumStringReferences = 512;
                private const int MinimumDuplicateReferences = 128;
                private const int MinimumDuplicateCharacters = 4096;
                private const int MaximumSeenOnceCandidates = 8192;
                private const int MinimumEarlyUniqueHeavyStringReferences = 16384;
                private const long MinimumDuplicateCharacterShareNumerator = 3L;
                private const long MinimumDuplicateCharacterShareDenominator = 5L;
                private readonly Dictionary<string, int> _indexes;

                private DirectSharedStringTable(Dictionary<string, int> indexes, string[] values, int totalStringReferences) {
                    _indexes = indexes;
                    Values = values;
                    TotalStringReferences = totalStringReferences;
                }

                internal string[] Values { get; }

                internal int TotalStringReferences { get; }

                internal bool TryGetIndex(string value, out int index) => _indexes.TryGetValue(value, out index);

                internal static DirectSharedStringTable? Create(DirectDataSetWorkbookModel model, IReadOnlyList<DirectColumnWritePlan> columnWritePlans, CancellationToken ct) {
                    if (!CanReachMinimumStringReferences(model, columnWritePlans)) {
                        return null;
                    }

                    var stringCounts = new Dictionary<string, int>(StringComparer.Ordinal);
                    int totalStringReferences = 0;
                    int duplicateReferences = 0;
                    int sharedValueCount = 0;
                    int seenOnceCount = 0;
                    long totalStringCharacters = 0L;
                    long duplicateCharacters = 0L;
                    bool canCancel = ct.CanBeCanceled;
                    for (int sheetIndex = 0; sheetIndex < model.Sheets.Count; sheetIndex++) {
                        var sheet = model.Sheets[sheetIndex];
                        if (sheet.Table.TryGetCellValueRows(out _)
                            || sheet.Table.TryGetExactDictionaryRows(out _)
                            || sheet.Table.TryGetDictionaryRows(out _)
                            || sheet.Table.TryGetLegacyDictionaryRows(out _)) {
                            return null;
                        }

                        int columnCount = sheet.Table.ColumnCount;
                        if (sheet.IncludeHeaders) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                NoteString(sheet.Table.GetColumnName(columnIndex), forceShared: true);
                            }
                        }

                        int[]? stringColumnIndexes = CreateSharedStringCandidateColumnIndexes(columnWritePlans[sheetIndex]);
                        if (stringColumnIndexes == null) {
                            continue;
                        }

                        int rowCount = sheet.Table.RowCount;
                        int stringColumnCount = stringColumnIndexes.Length;
                        if (sheet.Table.TryGetBufferedRows(out DirectBufferedRows bufferedRows)) {
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                object?[] bufferedRow = bufferedRows[rowIndex];
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (bufferedRow[columnIndex] is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else if (sheet.Table.HasSourceRows) {
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                DataRow sourceRow = sheet.Table.GetSourceRow(rowIndex)!;
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (sourceRow[columnIndex] is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else if (sheet.Table.TryGetExactDictionaryRows(out var exactDictionaryRows)) {
                            string[] columnNames = sheet.Table.CreateColumnNameArray();
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                Dictionary<string, object?> row = exactDictionaryRows[rowIndex];
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (row.TryGetValue(columnNames[columnIndex], out object? value)
                                        && value is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else if (sheet.Table.TryGetDictionaryRows(out var dictionaryRows)) {
                            string[] columnNames = sheet.Table.CreateColumnNameArray();
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                IReadOnlyDictionary<string, object?> row = dictionaryRows[rowIndex];
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (row.TryGetValue(columnNames[columnIndex], out object? value)
                                        && value is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else if (sheet.Table.TryGetLegacyDictionaryRows(out _)) {
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (sheet.Table.GetValue(rowIndex, columnIndex) is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else {
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (sheet.Table.GetValue(rowIndex, columnIndex) is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        }
                    }

                    if (totalStringReferences < MinimumStringReferences || sharedValueCount == 0) {
                        return null;
                    }

                    if (duplicateReferences < MinimumDuplicateReferences && duplicateCharacters < MinimumDuplicateCharacters) {
                        return null;
                    }

                    var indexes = new Dictionary<string, int>(sharedValueCount, StringComparer.Ordinal);
                    var values = new string[sharedValueCount];
                    int nextIndex = 0;
                    int sharedStringReferences = 0;
                    foreach (var entry in stringCounts) {
                        int referenceCount = entry.Value < 0 ? -entry.Value : entry.Value;
                        if (referenceCount == 1 && entry.Value > 0) {
                            continue;
                        }

                        CoerceValueHelper.ValidateSharedStringLength(entry.Key, "value");
                        values[nextIndex] = entry.Key;
                        indexes[entry.Key] = nextIndex;
                        sharedStringReferences += referenceCount;
                        nextIndex++;
                    }

                    return new DirectSharedStringTable(indexes, values, sharedStringReferences);

                    void NoteString(string text, bool forceShared = false) {
                        totalStringReferences++;
                        totalStringCharacters += text.Length;
                        if (stringCounts.TryGetValue(text, out int count)) {
                            if (count == 1) {
                                seenOnceCount--;
                                sharedValueCount++;
                                stringCounts[text] = 2;
                            } else if (count > 1) {
                                stringCounts[text] = count + 1;
                            } else {
                                stringCounts[text] = count - 1;
                            }

                            duplicateReferences++;
                            duplicateCharacters += text.Length;
                            return;
                        }

                        if (forceShared) {
                            stringCounts.Add(text, -1);
                            sharedValueCount++;
                            return;
                        }

                        if (seenOnceCount >= MaximumSeenOnceCandidates) {
                            return;
                        }

                        stringCounts.Add(text, 1);
                        seenOnceCount++;
                    }

                    bool ShouldAbandonSharedStrings() {
                        return totalStringReferences >= MinimumEarlyUniqueHeavyStringReferences
                            && seenOnceCount >= MaximumSeenOnceCandidates
                            && duplicateCharacters * MinimumDuplicateCharacterShareDenominator < totalStringCharacters * MinimumDuplicateCharacterShareNumerator;
                    }
                }

                private static bool CanReachMinimumStringReferences(DirectDataSetWorkbookModel model, IReadOnlyList<DirectColumnWritePlan> columnWritePlans) {
                    long possibleStringReferences = 0L;
                    for (int sheetIndex = 0; sheetIndex < model.Sheets.Count; sheetIndex++) {
                        var sheet = model.Sheets[sheetIndex];
                        int columnCount = sheet.Table.ColumnCount;
                        if (sheet.IncludeHeaders) {
                            possibleStringReferences += columnCount;
                        }

                        int[]? stringColumnIndexes = CreateSharedStringCandidateColumnIndexes(columnWritePlans[sheetIndex]);
                        if (stringColumnIndexes != null) {
                            int rowCount = sheet.Table.RowCount;
                            possibleStringReferences += (long)rowCount * stringColumnIndexes.Length;
                        }

                        if (possibleStringReferences >= MinimumStringReferences) {
                            return true;
                        }
                    }

                    return false;
                }

                private static int[]? CreateSharedStringCandidateColumnIndexes(DirectColumnWritePlan columnWritePlan) {
                    DirectCellValueKind[] kinds = columnWritePlan.CellValueKinds;
                    bool[]? valueStyleColumns = columnWritePlan.ValueStyleColumns;
                    int[]? indexes = null;
                    int count = 0;
                    for (int i = 0; i < kinds.Length; i++) {
                        bool canContainString = kinds[i] == DirectCellValueKind.String
                            || valueStyleColumns?[i] == true;
                        if (!canContainString) {
                            continue;
                        }

                        indexes ??= new int[kinds.Length];
                        indexes[count++] = i;
                    }

                    if (indexes == null) {
                        return null;
                    }

                    if (count == indexes.Length) {
                        return indexes;
                    }

                    Array.Resize(ref indexes, count);
                    return indexes;
                }
            }

            private static string? GetStyleAttribute(DirectCellValueKind cellValueKind, bool useCellValueNumberFormats) {
                switch (cellValueKind) {
                    case DirectCellValueKind.DateTime:
                    case DirectCellValueKind.DateTimeOffset:
                        return useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;
                    case DirectCellValueKind.TimeSpan:
                        return useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;
#if NET6_0_OR_GREATER
                    case DirectCellValueKind.DateOnly:
                        return useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;
                    case DirectCellValueKind.TimeOnly:
                        return useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;
#endif
                    default:
                        return null;
                }
            }
        }

    }
}
