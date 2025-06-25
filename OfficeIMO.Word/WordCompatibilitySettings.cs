using System;
using DocumentFormat.OpenXml;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    /// <summary>
    /// Specifies compatibility modes corresponding to different Word versions.
    /// </summary>
    public enum CompatibilityMode {
        /// <summary>
        /// No specific compatibility mode is applied.
        /// </summary>
        None = 0,
        /// <summary>
        /// Compatibility with Word 2003 (version 11).
        /// </summary>
        Word2003 = 11,
        /// <summary>
        /// Compatibility with Word 2007 (version 12).
        /// </summary>
        Word2007 = 12,
        /// <summary>
        /// Compatibility with Word 2010 (version 14).
        /// </summary>
        Word2010 = 14,
        /// <summary>
        /// Compatibility with Word 2013 (version 15).
        /// </summary>
        Word2013 = 15
    }

    /// <summary>
    /// Provides access to compatibility settings for the document and allows
    /// reading or changing the compatibility mode.
    /// </summary>
    public class WordCompatibilitySettings {
        private WordprocessingDocument _wordprocessingDocument;
        private WordDocument _document;

        /// <summary>
        /// Initializes a new instance of <see cref="WordCompatibilitySettings"/>
        /// for the specified document.
        /// </summary>
        /// <param name="document">Word document associated with the settings.</param>
        public WordCompatibilitySettings(WordDocument document) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;
            document.CompatibilitySettings = this;
        }

        /// <summary>
        /// Gets or sets compatibility mode of a Word Document
        /// </summary>
        public CompatibilityMode CompatibilityMode {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                var compatibility = settings.OfType<Compatibility>().FirstOrDefault();
                if (compatibility == null) {
                    return CompatibilityMode.None;
                }
                foreach (var setting in compatibility.OfType<CompatibilitySetting>()) {
                    if (setting.Name == CompatSettingNameValues.CompatibilityMode) {
                        return (CompatibilityMode)int.Parse(setting.Val);
                    }
                }

                return CompatibilityMode.None;
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                var compatibility = settings.OfType<Compatibility>().FirstOrDefault();
                if (compatibility == null) {
                    compatibility = new Compatibility();
                    settings.Append(compatibility);
                }

                foreach (var setting in compatibility.OfType<CompatibilitySetting>()) {
                    if (setting.Name == CompatSettingNameValues.CompatibilityMode) {
                        if (value == CompatibilityMode.None) {
                            setting.Remove();
                        } else {
                            setting.Val = ((int)value).ToString();
                            setting.Uri = "http://schemas.microsoft.com/office/word";
                        }
                        return;
                    }
                }
                compatibility.Append(new CompatibilitySetting() {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = ((int)value).ToString()
                });
            }
        }
    }
}
