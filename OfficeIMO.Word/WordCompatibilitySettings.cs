using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

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
        private readonly WordprocessingDocument _wordprocessingDocument;
        private readonly WordDocument _document;

        /// <summary>
        /// Initializes a new instance of <see cref="WordCompatibilitySettings"/>
        /// for the specified document.
        /// </summary>
        /// <param name="document">Word document associated with the settings.</param>
        public WordCompatibilitySettings(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument ?? throw new InvalidOperationException("The document does not contain an associated WordprocessingDocument instance.");
            document.CompatibilitySettings = this;
        }

        /// <summary>
        /// Gets or sets compatibility mode of a Word Document
        /// </summary>
        public CompatibilityMode CompatibilityMode {
            get {
                Settings settings = GetSettings();
                Compatibility? compatibility = settings.Elements<Compatibility>().FirstOrDefault();
                if (compatibility == null) {
                    return CompatibilityMode.None;
                }
                foreach (CompatibilitySetting setting in compatibility.Elements<CompatibilitySetting>()) {
                    if (setting.Name?.Value == CompatSettingNameValues.CompatibilityMode) {
                        string? valueText = setting.Val?.Value ?? setting.Val;
                        if (int.TryParse(valueText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int modeValue) && Enum.IsDefined(typeof(CompatibilityMode), modeValue)) {
                            return (CompatibilityMode)modeValue;
                        }
                        break;
                    }
                }

                return CompatibilityMode.None;
            }
            set {
                Settings settings = GetSettings();
                Compatibility? compatibility = settings.Elements<Compatibility>().FirstOrDefault();
                if (compatibility == null) {
                    compatibility = new Compatibility();
                    settings.Append(compatibility);
                }

                foreach (CompatibilitySetting setting in compatibility.Elements<CompatibilitySetting>()) {
                    if (setting.Name?.Value == CompatSettingNameValues.CompatibilityMode) {
                        if (value == CompatibilityMode.None) {
                            setting.Remove();
                        } else {
                            setting.Val = ((int)value).ToString(CultureInfo.InvariantCulture);
                            setting.Uri = "http://schemas.microsoft.com/office/word";
                        }

                        return;
                    }
                }

                if (value != CompatibilityMode.None) {
                    compatibility.Append(new CompatibilitySetting {
                        Name = CompatSettingNameValues.CompatibilityMode,
                        Uri = "http://schemas.microsoft.com/office/word",
                        Val = ((int)value).ToString(CultureInfo.InvariantCulture)
                    });
                }
            }
        }

        private Settings GetSettings() {
            MainDocumentPart mainPart = _wordprocessingDocument.MainDocumentPart ?? throw new InvalidOperationException("The document does not contain a main document part.");
            DocumentSettingsPart settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings ??= new Settings();
            return settingsPart.Settings;
        }
    }
}
