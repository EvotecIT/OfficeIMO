using System;
using System.Collections.Generic;
using System.IO;
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
            var names = MinimalVbaParser.GetModuleNames(stream);
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
    }
}
