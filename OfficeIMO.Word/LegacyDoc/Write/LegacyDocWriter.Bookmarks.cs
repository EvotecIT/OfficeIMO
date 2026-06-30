using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static byte[] CreateBookmarkNameTable(IReadOnlyList<LegacyDocWritableBookmark> bookmarks) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0xFFFF);
            WriteUInt16(stream, checked((ushort)bookmarks.Count));
            WriteUInt16(stream, 0);
            foreach (LegacyDocWritableBookmark bookmark in bookmarks) {
                WriteUInt16(stream, checked((ushort)bookmark.Name.Length));
                byte[] nameBytes = Encoding.Unicode.GetBytes(bookmark.Name);
                stream.Write(nameBytes, 0, nameBytes.Length);
            }

            return stream.ToArray();
        }

        private static byte[] CreateBookmarkStartPlc(IReadOnlyList<LegacyDocWritableBookmark> startOrderedBookmarks, IReadOnlyDictionary<string, int> endIndexes, int terminalCharacterPosition) {
            int count = startOrderedBookmarks.Count;
            var plc = new byte[((count + 1) * 4) + (count * 4)];
            for (int index = 0; index < count; index++) {
                WriteInt32(plc, index * 4, startOrderedBookmarks[index].StartCharacter);
            }

            WriteInt32(plc, count * 4, terminalCharacterPosition);
            int dataOffset = (count + 1) * 4;
            for (int index = 0; index < count; index++) {
                LegacyDocWritableBookmark bookmark = startOrderedBookmarks[index];
                WriteUInt16(plc, dataOffset + (index * 4), checked((ushort)endIndexes[bookmark.Id]));
                WriteUInt16(plc, dataOffset + (index * 4) + 2, 0);
            }

            return plc;
        }

        private static byte[] CreateBookmarkEndPlc(IReadOnlyList<LegacyDocWritableBookmark> endOrderedBookmarks, int terminalCharacterPosition) {
            int count = endOrderedBookmarks.Count;
            var plc = new byte[(count + 1) * 4];
            for (int index = 0; index < count; index++) {
                WriteInt32(plc, index * 4, endOrderedBookmarks[index].EndCharacter);
            }

            WriteInt32(plc, count * 4, terminalCharacterPosition);
            return plc;
        }

        private sealed class LegacyDocWritableBookmarksBuilder {
            private readonly Dictionary<string, LegacyDocWritableBookmarkStart> _starts = new(StringComparer.Ordinal);
            private readonly List<LegacyDocWritableBookmark> _bookmarks = new();
            private readonly HashSet<string> _names = new(StringComparer.Ordinal);
            private readonly HashSet<string> _ids = new(StringComparer.Ordinal);

            internal void AddStart(BookmarkStart bookmarkStart, int startCharacter) {
                string id = ReadBookmarkId(bookmarkStart.Id?.Value, "start");
                string name = ReadBookmarkName(bookmarkStart.Name?.Value);
                if (!_names.Add(name)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate bookmark name '{name}'.");
                }

                if (!_ids.Add(id)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate bookmark id '{id}'.");
                }

                _starts.Add(id, new LegacyDocWritableBookmarkStart(id, name, startCharacter));
            }

            internal void AddRange(LegacyDocWritableBookmarks bookmarks, int characterOffset) {
                foreach (LegacyDocWritableBookmark bookmark in bookmarks.StartOrderedBookmarks) {
                    AddCompleted(
                        bookmark.Id,
                        bookmark.Name,
                        bookmark.StartCharacter + characterOffset,
                        bookmark.EndCharacter + characterOffset);
                }
            }

            internal void AddEnd(BookmarkEnd bookmarkEnd, int endCharacter) {
                string id = ReadBookmarkId(bookmarkEnd.Id?.Value, "end");
                if (!_starts.TryGetValue(id, out LegacyDocWritableBookmarkStart start)) {
                    throw new NotSupportedException($"Native DOC saving cannot write bookmark end id '{id}' because the matching bookmark start was not found in the same supported story.");
                }

                if (endCharacter < start.StartCharacter) {
                    throw new NotSupportedException($"Native DOC saving cannot write bookmark '{start.Name}' because its end is before its start.");
                }

                _starts.Remove(id);
                _bookmarks.Add(new LegacyDocWritableBookmark(id, start.Name, start.StartCharacter, endCharacter));
            }

            private void AddCompleted(string id, string name, int startCharacter, int endCharacter) {
                if (!_names.Add(name)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate bookmark name '{name}'.");
                }

                if (!_ids.Add(id)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate bookmark id '{id}'.");
                }

                if (endCharacter < startCharacter) {
                    throw new NotSupportedException($"Native DOC saving cannot write bookmark '{name}' because its end is before its start.");
                }

                _bookmarks.Add(new LegacyDocWritableBookmark(id, name, startCharacter, endCharacter));
            }

            internal LegacyDocWritableBookmarks Create() {
                if (_starts.Count != 0) {
                    string names = string.Join(", ", _starts.Values.Select(start => start.Name).Take(5));
                    throw new NotSupportedException($"Native DOC saving cannot write bookmarks without matching end markers. Unclosed bookmark names: {names}.");
                }

                if (_bookmarks.Count == 0) {
                    return LegacyDocWritableBookmarks.Empty;
                }

                if (_bookmarks.Count > 0x3FFB) {
                    throw new NotSupportedException("Native DOC saving supports at most 0x3FFB standard bookmarks.");
                }

                LegacyDocWritableBookmark[] startOrdered = _bookmarks
                    .OrderBy(bookmark => bookmark.StartCharacter)
                    .ThenBy(bookmark => bookmark.EndCharacter)
                    .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)
                    .ToArray();
                LegacyDocWritableBookmark[] endOrdered = _bookmarks
                    .OrderBy(bookmark => bookmark.EndCharacter)
                    .ThenBy(bookmark => bookmark.StartCharacter)
                    .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)
                    .ToArray();
                Dictionary<string, int> endIndexes = endOrdered
                    .Select((bookmark, index) => new { bookmark.Id, index })
                    .ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);

                return new LegacyDocWritableBookmarks(
                    CreateBookmarkNameTable(startOrdered),
                    CreateBookmarkStartPlc(startOrdered, endIndexes, terminalCharacterPosition: 0),
                    CreateBookmarkEndPlc(endOrdered, terminalCharacterPosition: 0),
                    startOrdered,
                    endOrdered,
                    endIndexes);
            }

            private static string ReadBookmarkId(string? id, string markerKind) {
                if (string.IsNullOrWhiteSpace(id)) {
                    throw new NotSupportedException($"Native DOC saving cannot write a bookmark {markerKind} marker without an id.");
                }

                return id!;
            }

            private static string ReadBookmarkName(string? name) {
                if (string.IsNullOrWhiteSpace(name)) {
                    throw new NotSupportedException("Native DOC saving cannot write a bookmark without a name.");
                }

                if (name!.Length > 40) {
                    throw new NotSupportedException($"Native DOC saving cannot write bookmark '{name}' because Word 97-2003 bookmark names are limited to 40 characters.");
                }

                return name;
            }
        }

        private readonly struct LegacyDocWritableBookmarkStart {
            internal LegacyDocWritableBookmarkStart(string id, string name, int startCharacter) {
                Id = id;
                Name = name;
                StartCharacter = startCharacter;
            }

            internal string Id { get; }

            internal string Name { get; }

            internal int StartCharacter { get; }
        }

        private readonly struct LegacyDocWritableBookmark {
            internal LegacyDocWritableBookmark(string id, string name, int startCharacter, int endCharacter) {
                Id = id;
                Name = name;
                StartCharacter = startCharacter;
                EndCharacter = endCharacter;
            }

            internal string Id { get; }

            internal string Name { get; }

            internal int StartCharacter { get; }

            internal int EndCharacter { get; }
        }

        private readonly struct LegacyDocWritableBookmarks {
            internal static LegacyDocWritableBookmarks Empty { get; } = new(Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<byte>());

            internal LegacyDocWritableBookmarks(byte[] sttbfBkmk, byte[] plcfBkf, byte[] plcfBkl)
                : this(sttbfBkmk, plcfBkf, plcfBkl, Array.Empty<LegacyDocWritableBookmark>(), Array.Empty<LegacyDocWritableBookmark>(), new Dictionary<string, int>(StringComparer.Ordinal)) {
            }

            internal LegacyDocWritableBookmarks(
                byte[] sttbfBkmk,
                byte[] plcfBkf,
                byte[] plcfBkl,
                IReadOnlyList<LegacyDocWritableBookmark> startOrderedBookmarks,
                IReadOnlyList<LegacyDocWritableBookmark> endOrderedBookmarks,
                IReadOnlyDictionary<string, int> endIndexes) {
                SttbfBkmk = sttbfBkmk;
                PlcfBkf = plcfBkf;
                PlcfBkl = plcfBkl;
                StartOrderedBookmarks = startOrderedBookmarks;
                EndOrderedBookmarks = endOrderedBookmarks;
                EndIndexes = endIndexes;
            }

            internal byte[] SttbfBkmk { get; }

            internal byte[] PlcfBkf { get; }

            internal byte[] PlcfBkl { get; }

            internal IReadOnlyList<LegacyDocWritableBookmark> StartOrderedBookmarks { get; }

            internal IReadOnlyList<LegacyDocWritableBookmark> EndOrderedBookmarks { get; }

            internal IReadOnlyDictionary<string, int> EndIndexes { get; }

            internal LegacyDocWritableBookmarks WithTerminalCharacterPosition(int terminalCharacterPosition) {
                if (StartOrderedBookmarks.Count == 0) {
                    return this;
                }

                return new LegacyDocWritableBookmarks(
                    SttbfBkmk,
                    CreateBookmarkStartPlc(StartOrderedBookmarks, EndIndexes, terminalCharacterPosition),
                    CreateBookmarkEndPlc(EndOrderedBookmarks, terminalCharacterPosition),
                    StartOrderedBookmarks,
                    EndOrderedBookmarks,
                    EndIndexes);
            }
        }
    }
}
