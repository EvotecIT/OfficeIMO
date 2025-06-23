using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a single macro module within a document.
    /// </summary>
    /// <remarks>
    /// Instances are returned by <see cref="WordDocument.Macros"/> and can be
    /// removed individually using <see cref="Remove"/>.
    /// </remarks>
    public class WordMacro {
        private readonly WordDocument _document;

        /// <summary>
        /// Gets the macro module name.
        /// </summary>
        public string Name { get; }

        internal WordMacro(WordDocument document, string name) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        /// <summary>
        /// Removes this macro module from the document.
        /// </summary>
        public void Remove() {
            WordMacro.RemoveMacro(_document, Name);
        }

        internal static IReadOnlyList<WordMacro> GetMacros(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (!document.HasMacros) return new List<WordMacro>();

            var vbaPart = document._wordprocessingDocument.MainDocumentPart.VbaProjectPart;
            using var stream = vbaPart.GetStream();
            var names = Parser.GetModuleNames(stream);
            var modules = new List<WordMacro>(names.Count);
            foreach (var name in names) {
                modules.Add(new WordMacro(document, name));
            }
            return modules;
        }

        internal static void AddMacro(WordDocument document, string filePath) {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("File doesn't exist", filePath);

            AddMacro(document, File.ReadAllBytes(filePath));
        }

        internal static void AddMacro(WordDocument document, byte[] data) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (data == null || data.Length == 0) throw new ArgumentNullException(nameof(data));

            var main = document._wordprocessingDocument.MainDocumentPart;
            if (main.VbaProjectPart != null) {
                main.DeletePart(main.VbaProjectPart);
            }
            var vbaPart = main.AddNewPart<VbaProjectPart>();
            using var stream = new MemoryStream(data);
            vbaPart.FeedData(stream);
        }

        internal static byte[] ExtractMacros(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var vbaPart = document._wordprocessingDocument.MainDocumentPart.VbaProjectPart;
            if (vbaPart == null) return null;
            using var ms = new MemoryStream();
            using var partStream = vbaPart.GetStream();
            partStream.CopyTo(ms);
            return ms.ToArray();
        }

        internal static void SaveMacros(WordDocument document, string filePath) {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentNullException(nameof(filePath));
            var data = ExtractMacros(document);
            if (data == null) return;
            File.WriteAllBytes(filePath, data);
        }

        internal static void RemoveMacro(WordDocument document, string name) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (!document.HasMacros) return;
            RemoveMacros(document);
        }

        internal static void RemoveMacros(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var main = document._wordprocessingDocument.MainDocumentPart;
            if (main.VbaProjectPart != null) {
                main.DeletePart(main.VbaProjectPart);
            }
        }

        private static class Parser {
            private const int EndOfChain = unchecked((int)0xFFFFFFFE);

            private class DirEntry {
                public string Name = string.Empty;
                public byte Type;
                public int Left;
                public int Right;
                public int Child;
                public int StartSector;
                public long Size;
                public int Parent = -1;
            }

            internal static IReadOnlyList<string> GetModuleNames(Stream stream) {
                var modules = new List<string>();
                if (stream == null || !stream.CanRead) return modules;

                using var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true);
                byte[] header = reader.ReadBytes(512);
                if (header.Length < 512) return modules;
                const ulong Signature = 0xE11AB1A1E011CFD0UL;
                if (BitConverter.ToUInt64(header, 0) != Signature) return modules;

                ushort sectorShift = BitConverter.ToUInt16(header, 0x1E);
                int sectorSize = 1 << sectorShift;
                int dirStart = BitConverter.ToInt32(header, 0x30);

                var fatSectors = new List<int>();
                for (int i = 0; i < 109; i++) {
                    int s = BitConverter.ToInt32(header, 0x4C + i * 4);
                    if (s >= 0) fatSectors.Add(s);
                }
                var fat = ReadFat(stream, reader, fatSectors, sectorSize);
                byte[] dirData = ReadChain(stream, reader, dirStart, sectorSize, fat);
                if (dirData.Length == 0) return modules;

                var entries = ParseDirectory(dirData);
                var queue = new Queue<(int id, int parent)>();
                var visited = new HashSet<int>();
                queue.Enqueue((0, -1));
                int vbaId = -1;
                while (queue.Count > 0) {
                    var (id, parent) = queue.Dequeue();
                    if (id < 0 || id >= entries.Count) continue;
                    if (!visited.Add(id)) continue;
                    var e = entries[id];
                    e.Parent = parent;
                    if (e.Name == "VBA" && e.Type == 1) vbaId = id;
                    if (e.Left >= 0) queue.Enqueue((e.Left, parent));
                    if (e.Right >= 0) queue.Enqueue((e.Right, parent));
                    if (e.Child >= 0) queue.Enqueue((e.Child, id));
                }
                if (vbaId >= 0) {
                    foreach (var e in entries) {
                        if (e.Parent == vbaId && e.Type == 2 && e.Name != "dir" && e.Name != "_VBA_PROJECT" && e.Name != "PROJECT") {
                            modules.Add(e.Name);
                        }
                    }
                }
                return modules;
            }

            private static List<int> ReadFat(Stream stream, BinaryReader reader, List<int> fatSectors, int sectorSize) {
                var fat = new List<int>();
                int perSector = sectorSize / 4;
                foreach (int sec in fatSectors) {
                    if (sec < 0) continue;
                    stream.Position = (long)(sec + 1) * sectorSize;
                    for (int i = 0; i < perSector; i++) fat.Add(reader.ReadInt32());
                }
                return fat;
            }

            private static byte[] ReadChain(Stream stream, BinaryReader reader, int start, int sectorSize, List<int> fat) {
                if (start < 0) return Array.Empty<byte>();
                using var ms = new MemoryStream();
                int sector = start;
                while (sector != EndOfChain && sector >= 0 && sector < fat.Count) {
                    stream.Position = (long)(sector + 1) * sectorSize;
                    ms.Write(reader.ReadBytes(sectorSize));
                    sector = fat[sector];
                }
                return ms.ToArray();
            }

            private static List<DirEntry> ParseDirectory(byte[] data) {
                var list = new List<DirEntry>(data.Length / 128);
                for (int offset = 0; offset + 128 <= data.Length; offset += 128) {
                    ushort nameLen = BitConverter.ToUInt16(data, offset + 64);
                    string name = nameLen > 0 ? Encoding.Unicode.GetString(data, offset, nameLen - 2) : string.Empty;
                    byte type = data[offset + 66];
                    int left = BitConverter.ToInt32(data, offset + 68);
                    int right = BitConverter.ToInt32(data, offset + 72);
                    int child = BitConverter.ToInt32(data, offset + 76);
                    int start = BitConverter.ToInt32(data, offset + 116);
                    long size = BitConverter.ToInt64(data, offset + 120);
                    list.Add(new DirEntry { Name = name, Type = type, Left = left, Right = right, Child = child, StartSector = start, Size = size });
                }
                return list;
            }
        }
    }
}
