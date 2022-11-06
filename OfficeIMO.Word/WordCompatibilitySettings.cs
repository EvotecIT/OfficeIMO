using System;
using DocumentFormat.OpenXml;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    public enum CompatibilityMode {
        None = 0,
        Word2003 = 11,
        Word2007 = 12,
        Word2010 = 14,
        Word2013 = 15
    }

    public class WordCompatibilitySettings {
        private WordprocessingDocument _wordprocessingDocument;
        private WordDocument _document;

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
