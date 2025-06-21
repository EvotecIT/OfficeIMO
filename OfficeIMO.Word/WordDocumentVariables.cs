using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides access to document variables stored in the document settings.
    /// </summary>
    public class WordDocumentVariables {
        private readonly WordprocessingDocument _wordprocessingDocument;
        private readonly WordDocument _document;
        private DocumentVariables _variables;

        /// <summary>
        /// Initializes a new instance of <see cref="WordDocumentVariables"/> and loads or creates document variables.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="create">When set to <c>true</c> variables are written to the document.</param>
        public WordDocumentVariables(WordDocument document, bool? create = null) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;

            if (create == true) {
                CreateDocumentVariables();
            } else {
                LoadDocumentVariables();
            }
        }

        private void LoadDocumentVariables() {
            var settingsPart = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart;
            if (settingsPart?.Settings != null) {
                _variables = settingsPart.Settings.GetFirstChild<DocumentVariables>();
                if (_variables != null) {
                    foreach (var variable in _variables.Elements<DocumentVariable>()) {
                        _document.DocumentVariables[variable.Name] = variable.Val;
                    }
                }
            }
        }

        private void CreateDocumentVariables() {
            var settingsPart = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart;
            if (settingsPart == null) {
                if (_document.FileOpenAccess == FileAccess.Read) {
                    throw new ArgumentException("Document is read only!");
                }
                settingsPart = _wordprocessingDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();
            }

            var settings = settingsPart.Settings;
            _variables = settings.GetFirstChild<DocumentVariables>();

            if (_document.DocumentVariables.Count == 0) {
                _variables?.Remove();
                return;
            }

            if (_variables == null) {
                _variables = new DocumentVariables();
                settings.Append(_variables);
            }

            // remove variables not present in the dictionary
            var toRemove = _variables.Elements<DocumentVariable>()
                .Where(v => !_document.DocumentVariables.ContainsKey(v.Name))
                .ToList();
            foreach (var variable in toRemove) {
                variable.Remove();
            }

            foreach (var pair in _document.DocumentVariables) {
                var existing = _variables.Elements<DocumentVariable>()
                    .FirstOrDefault(v => v.Name == pair.Key);
                if (existing != null) {
                    existing.Val = pair.Value;
                } else {
                    _variables.Append(new DocumentVariable { Name = pair.Key, Val = pair.Value });
                }
            }
        }
    }
}
