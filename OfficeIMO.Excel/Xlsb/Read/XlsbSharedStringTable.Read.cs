using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using System.Text;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbSharedStringTable {
        private const int BrtSstItem = 19;
        private const int BrtBeginSst = 159;
        private const int BrtEndSst = 160;

        /// <summary>
        /// Takes ownership of the pooled part, validating every record before publishing any
        /// text. Optimization thresholds select representation and never relax configured limits.
        /// </summary>
        internal static XlsbSharedStringTable Read(
            XlsbPooledPartStream part, string partName, XlsbImportOptions limits,
            XlsbRecordReadBudget recordBudget, int maxItemCharacters, long maxCharacters,
            Action checkCancellation) {
            bool retained = false;
            try {
                var records = new XlsbRecordSliceReader(part.Buffer, limits.MaxRecordBytes, recordBudget, length: part.DataLength);
                bool indexed = part.DataLength <= MaximumIndexedPartBytes && part.Buffer.Length <= MaximumIndexedPartBytes;
                List<TextEntry>? entries = indexed ? new List<TextEntry>() : null;
                List<string>? values = indexed ? null : new List<string>();
                long totalCharacters = 0;
                bool began = false;
                bool ended = false;
                uint declaredUniqueCount = 0;
                while (records.TryRead(out XlsbRecordSlice record)) {
                    checkCancellation();
                    if (record.Type == BrtBeginSst) {
                        if (began || ended || record.RecordOffset != 0 || record.Size < 8) {
                            throw new InvalidDataException($"The BrtBeginSst record in '{partName}' is truncated or misplaced.");
                        }
                        var header = record.CreateCursor();
                        uint declaredTotalCount = header.ReadUInt32();
                        declaredUniqueCount = header.ReadUInt32();
                        if (declaredTotalCount < declaredUniqueCount) {
                            throw new InvalidDataException($"The XLSB shared-string part '{partName}' declares fewer total strings than unique strings.");
                        }
                        if (declaredUniqueCount > limits.MaxSharedStrings) {
                            throw new InvalidDataException($"The XLSB shared-string table declares {declaredUniqueCount} unique items, exceeding the configured limit of {limits.MaxSharedStrings} items.");
                        }
                        began = true;
                    } else if (record.Type == BrtEndSst) {
                        if (!began || ended || record.Size != 0) {
                            throw new InvalidDataException($"The XLSB shared-string part '{partName}' contains an invalid BrtEndSst boundary.");
                        }
                        ended = true;
                    } else if (record.Type == BrtSstItem) {
                        if (!began || ended) {
                            throw new InvalidDataException($"The XLSB shared-string part '{partName}' contains an item outside its boundary records.");
                        }
                        int count = entries?.Count ?? values!.Count;
                        if (count >= limits.MaxSharedStrings) {
                            throw new InvalidDataException($"The XLSB shared-string table exceeds the configured limit of {limits.MaxSharedStrings} items.");
                        }
                        var cursor = record.CreateCursor();
                        cursor.ReadByte();
                        int offset = cursor.Position + sizeof(uint);
                        string? value = entries == null ? cursor.ReadWideString(maxItemCharacters) : null;
                        int characterCount = value?.Length ?? cursor.ValidateWideString(maxItemCharacters);
                        int byteCount = cursor.Position - offset;
                        if (totalCharacters > maxCharacters - characterCount) {
                            throw new InvalidDataException($"The XLSB shared-string table exceeds the configured aggregate limit of {maxCharacters} characters.");
                        }
                        totalCharacters += characterCount;
                        if (entries != null && count == MaximumIndexedItems) {
                            // The count is observed and validated, never an unchecked declaration.
                            values = new List<string>(entries.Count);
                            foreach (TextEntry entry in entries) {
                                checkCancellation();
                                values.Add(Encoding.Unicode.GetString(part.Buffer, entry.Offset, entry.ByteCount));
                            }
                            entries = null;
                        }
                        if (entries != null) entries.Add(new TextEntry(offset, byteCount));
                        else values!.Add(value ?? Encoding.Unicode.GetString(part.Buffer, offset, byteCount));
                    } else if (!began || ended) {
                        throw new InvalidDataException($"The XLSB shared-string part '{partName}' contains records outside its boundary records.");
                    }
                }
                if (!began || !ended) {
                    throw new InvalidDataException($"The XLSB shared-string part '{partName}' is missing its boundary records.");
                }
                int actualCount = entries?.Count ?? values!.Count;
                if ((uint)actualCount != declaredUniqueCount) {
                    throw new InvalidDataException($"The XLSB shared-string part '{partName}' declares {declaredUniqueCount} unique items but contains {actualCount}.");
                }
                if (entries != null) {
                    var table = new XlsbSharedStringTable(part, entries);
                    retained = true;
                    return table;
                }
                return new XlsbSharedStringTable(values!);
            } finally {
                if (!retained) part.Dispose();
            }
        }
    }
}
