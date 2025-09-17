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
        private DocumentVariables? _variables;

        /// <summary>
        /// Initializes a new instance of <see cref="WordDocumentVariables"/> and loads or creates document variables.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="create">When set to <c>true</c> variables are written to the document.</param>
        public WordDocumentVariables(WordDocument document, bool? create = null) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _wordprocessingDocument = document._wordprocessingDocument ?? throw new ArgumentNullException(nameof(document._wordprocessingDocument));

            if (create == true) {
                CreateDocumentVariables();
            } else {
                LoadDocumentVariables();
            }
        }

        private void LoadDocumentVariables() {
            var mainDocumentPart = _wordprocessingDocument.MainDocumentPart;
            if (mainDocumentPart == null) {
                return;
            }

            var settingsPart = mainDocumentPart.DocumentSettingsPart;
            var settings = settingsPart?.Settings;
            if (settings == null) {
                return;
            }

            _variables = settings.GetFirstChild<DocumentVariables>();
            if (_variables == null) {
                return;
            }

            foreach (var variable in _variables.Elements<DocumentVariable>()) {
                var name = variable.Name?.Value;
                if (string.IsNullOrEmpty(name)) {
                    continue;
                }

                var value = variable.Val?.Value ?? string.Empty;
                _document.DocumentVariables[name!] = value;
            }
        }

        private void CreateDocumentVariables() {
            var mainDocumentPart = _wordprocessingDocument.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            var settingsPart = mainDocumentPart.DocumentSettingsPart;
            if (settingsPart == null) {
                if (_document.FileOpenAccess == FileAccess.Read) {
                    throw new ArgumentException("Document is read only!");
                }
                settingsPart = mainDocumentPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();
            }

            var settings = settingsPart.Settings ??= new Settings();
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
                .Where(v => {
                    var name = v.Name?.Value;
                    return string.IsNullOrEmpty(name) || (name != null && !_document.DocumentVariables.ContainsKey(name));
                })
                .ToList();
            foreach (var variable in toRemove) {
                variable.Remove();
            }

            foreach (var pair in _document.DocumentVariables) {
                var existing = _variables.Elements<DocumentVariable>()
                    .FirstOrDefault(v => string.Equals(v.Name?.Value, pair.Key, StringComparison.Ordinal));
                if (existing != null) {
                    existing.Val = pair.Value;
                } else {
                    _variables.Append(new DocumentVariable { Name = pair.Key, Val = pair.Value });
                }
            }
        }
    }
}
