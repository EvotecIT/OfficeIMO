using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        /// <summary>Story-relative MS-DOC Plcfld records for the writer's supported field instructions.</summary>
        private sealed class LegacyDocWritableFieldTables {
            private static readonly int[] FibOffsets = { 0x11A, 0x122, 0x12A, 0x132, 0x21A };
            private readonly byte[][] _tables;

            internal LegacyDocWritableFieldTables(string body, IReadOnlyList<LegacyDocWritableRun> bodyRuns,
                string headers, IReadOnlyList<LegacyDocWritableRun> headerRuns,
                string footnotes, IReadOnlyList<LegacyDocWritableRun> footnoteRuns,
                string comments, IReadOnlyList<LegacyDocWritableRun> commentRuns,
                string endnotes, IReadOnlyList<LegacyDocWritableRun> endnoteRuns) {
                _tables = new[] { CreateFieldTable(body, bodyRuns), CreateFieldTable(headers, headerRuns),
                    CreateFieldTable(footnotes, footnoteRuns), CreateFieldTable(comments, commentRuns), CreateFieldTable(endnotes, endnoteRuns) };
                Length = _tables.Sum(table => table.Length);
            }

            internal int Length { get; }

            internal void WriteFibRecords(byte[] stream, int start) {
                for (int index = 0; index < _tables.Length; index++) {
                    byte[] table = _tables[index];
                    WriteInt32(stream, FibOffsets[index], table.Length == 0 ? 0 : start);
                    WriteInt32(stream, FibOffsets[index] + 4, table.Length);
                    start = checked(start + table.Length);
                }
            }

            internal void WriteTableBytes(byte[] destination, int start) {
                foreach (byte[] table in _tables) {
                    Buffer.BlockCopy(table, 0, destination, start, table.Length);
                    start = checked(start + table.Length);
                }
            }

            private static byte[] CreateFieldTable(string text, IReadOnlyList<LegacyDocWritableRun> runs) {
                // Only control characters carrying sprmCFSpec are fields. Plain
                // text with the same code point must not create a field record.
                var positions = new SortedSet<int>();
                foreach (LegacyDocWritableRun run in runs) {
                    if (!run.Formatting.Special) continue;
                    int end = checked(run.StartCharacter + run.Length);
                    for (int position = run.StartCharacter; position < end; position++) {
                        char character = text[position];
                        if (character == LegacyDocField.Begin || character == LegacyDocField.Separator || character == LegacyDocField.End)
                            positions.Add(position);
                    }
                }
                if (positions.Count == 0) return Array.Empty<byte>();
                var records = new List<(int Position, byte Character, byte Flags)>();
                var stack = new Stack<FieldFrame>();
                foreach (int position in positions) {
                    char character = text[position];
                    byte flags = 0;
                    if (character == LegacyDocField.Begin) {
                        flags = GetFieldType(text, position + 1);
                        stack.Push(new FieldFrame(stack.Count > 0));
                    } else if (character == LegacyDocField.Separator) {
                        if (stack.Count == 0 || stack.Peek().HasSeparator) throw InvalidFieldSequence();
                        stack.Peek().HasSeparator = true;
                    } else {
                        if (stack.Count == 0) throw InvalidFieldSequence();
                        FieldFrame frame = stack.Pop();
                        flags = (byte)((frame.HasSeparator ? 0x80 : 0) | (frame.Nested ? 0x40 : 0));
                    }
                    records.Add((position, (byte)character, flags));
                }
                if (stack.Count != 0) throw InvalidFieldSequence();
                int propertyStart = checked((records.Count + 1) * sizeof(int));
                byte[] result = new byte[checked(propertyStart + records.Count * 2)];
                for (int index = 0; index < records.Count; index++) {
                    var record = records[index];
                    WriteInt32(result, index * sizeof(int), record.Position);
                    result[propertyStart + index * 2] = record.Character;
                    result[propertyStart + index * 2 + 1] = record.Flags;
                }
                WriteInt32(result, records.Count * sizeof(int), text.Length);
                return result;
            }

            private static byte GetFieldType(string text, int start) {
                while (start < text.Length && char.IsWhiteSpace(text[start])) start++;
                int end = start;
                while (end < text.Length && char.IsLetter(text[end])) end++;
                return text.Substring(start, end - start).ToUpperInvariant() switch {
                    "TITLE" => 0x0F, "SUBJECT" => 0x10, "AUTHOR" => 0x11, "KEYWORDS" => 0x12,
                    "COMMENTS" => 0x13, "LASTSAVEDBY" => 0x14, "CREATEDATE" => 0x15,
                    "SAVEDATE" => 0x16, "PRINTDATE" => 0x17, "REVNUM" => 0x18, "NUMPAGES" => 0x1A,
                    "DATE" => 0x1F, "TIME" => 0x20, "PAGE" => 0x21,
                    "EQ" => 0x31, "DOCPROPERTY" => 0x55, "HYPERLINK" => 0x58,
                    _ => 0x01 // MS-DOC: the instruction has not been parsed.
                };
            }

            private static NotSupportedException InvalidFieldSequence() =>
                new("Native DOC saving requires balanced field characters with at most one result separator per field.");

            private sealed class FieldFrame {
                internal FieldFrame(bool nested) { Nested = nested; }
                internal bool Nested { get; }
                internal bool HasSeparator { get; set; }
            }
        }
    }
}
