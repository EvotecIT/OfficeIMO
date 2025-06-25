using System;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    /// <summary>
    /// Adds support for VBA macros.
    /// </summary>
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
                return WordMacro.GetMacros(this);
            }
        }

        /// <summary>
        /// Adds a VBA project to the document.
        /// </summary>
        /// <param name="filePath">Path to a <c>vbaProject.bin</c> file.</param>
        public void AddMacro(string filePath) {
            WordMacro.AddMacro(this, filePath);
        }

        /// <summary>
        /// Adds a VBA project to the document from bytes.
        /// </summary>
        /// <param name="data">VBA project data.</param>
        public void AddMacro(byte[] data) {
            WordMacro.AddMacro(this, data);
        }

        /// <summary>
        /// Extracts the VBA project as a byte array.
        /// </summary>
        /// <returns>Byte array with macro content or null when no macros are present.</returns>
        public byte[] ExtractMacros() {
            return WordMacro.ExtractMacros(this);
        }

        /// <summary>
        /// Saves the VBA project to a file.
        /// </summary>
        /// <param name="filePath">Destination path.</param>
        public void SaveMacros(string filePath) {
            WordMacro.SaveMacros(this, filePath);
        }

        /// <summary>
        /// Removes a single macro module from the document.
        /// </summary>
        /// <param name="name">Module name to remove.</param>
        public void RemoveMacro(string name) {
            WordMacro.RemoveMacro(this, name);
        }

        /// <summary>
        /// Removes the VBA project from the document.
        /// </summary>
        public void RemoveMacros() {
            WordMacro.RemoveMacros(this);
        }
    }
}
