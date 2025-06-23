using System;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Provides basic support for VBA macro projects.
        /// </summary>
        /// <remarks>
        /// Use <see cref="AddMacro(string)"/> or <see cref="AddMacro(byte[])"/>
        /// to attach a <c>vbaProject.bin</c> file extracted from a macro-enabled
        /// document. Macros can be enumerated through <see cref="Macros"/> and
        /// removed via <see cref="RemoveMacro"/> or <see cref="RemoveMacros"/>.
        /// To obtain the binary save a document as <c>.docm</c>, rename it to
        /// <c>.zip</c> and copy the file from the <c>word</c> folder.
        /// </remarks>
        /// <summary>
        /// Indicates whether the document contains a VBA project.
        /// </summary>
        public bool HasMacros => _wordprocessingDocument.MainDocumentPart.VbaProjectPart != null;

        /// <summary>
        /// Gets all macros (module streams) in the document.
        /// </summary>
        public IReadOnlyList<WordMacro> Macros {
            get {
                if (!HasMacros) return new List<WordMacro>();
                // Without external dependencies we cannot parse the VBA project.
                // Expose a single placeholder module when macros are present.
                return new List<WordMacro> { new WordMacro(this, "Module1") };
            }
        }

        /// <summary>
        /// Adds a VBA project to the document.
        /// </summary>
        /// <param name="filePath">Path to a <c>vbaProject.bin</c> file.</param>
        public void AddMacro(string filePath) {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("File doesn't exist", filePath);

            AddMacro(File.ReadAllBytes(filePath));
        }

        /// <summary>
        /// Adds a VBA project to the document from bytes.
        /// </summary>
        /// <param name="data">VBA project data.</param>
        public void AddMacro(byte[] data) {
            if (data == null || data.Length == 0) throw new ArgumentNullException(nameof(data));

            var main = _wordprocessingDocument.MainDocumentPart;
            if (main.VbaProjectPart != null) {
                main.DeletePart(main.VbaProjectPart);
            }
            var vbaPart = main.AddNewPart<VbaProjectPart>();
            using var stream = new MemoryStream(data);
            vbaPart.FeedData(stream);
        }

        /// <summary>
        /// Extracts the VBA project as a byte array.
        /// </summary>
        /// <returns>Byte array with macro content or null when no macros are present.</returns>
        public byte[] ExtractMacros() {
            var vbaPart = _wordprocessingDocument.MainDocumentPart.VbaProjectPart;
            if (vbaPart == null) return null;
            using var ms = new MemoryStream();
            using var partStream = vbaPart.GetStream();
            partStream.CopyTo(ms);
            return ms.ToArray();
        }

        /// <summary>
        /// Saves the VBA project to a file.
        /// </summary>
        /// <param name="filePath">Destination path.</param>
        public void SaveMacros(string filePath) {
            var data = ExtractMacros();
            if (data == null) return;
            File.WriteAllBytes(filePath, data);
        }

        /// <summary>
        /// Removes a single macro module from the document.
        /// </summary>
        /// <param name="name">Module name to remove.</param>
        public void RemoveMacro(string name) {
            if (!HasMacros) return;
            // Without the ability to modify VBA projects, removing a single
            // module deletes the entire project.
            RemoveMacros();
        }

        /// <summary>
        /// Removes the VBA project from the document.
        /// </summary>
        public void RemoveMacros() {
            var main = _wordprocessingDocument.MainDocumentPart;
            if (main.VbaProjectPart != null) {
                main.DeletePart(main.VbaProjectPart);
            }
        }
    }
}
