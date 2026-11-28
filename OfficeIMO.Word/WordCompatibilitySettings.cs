using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word {

    /// <summary>
    /// Specifies compatibility modes corresponding to different Word versions.
    /// </summary>
    public enum WordCompatibilityMode {
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
    /// reading or changing the compatibility mode and layout options.
    /// </summary>
    public partial class WordCompatibilitySettings {
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private const string CompatibilityModeName = "compatibilityMode";
        private const string CompatibilityModeUri = "http://schemas.microsoft.com/office/word";
        private readonly WordDocument _document;

        /// <summary>
        /// Initializes a new instance of <see cref="WordCompatibilitySettings"/>.
        /// </summary>
        /// <param name="document">Word document associated with the settings.</param>
        public WordCompatibilitySettings(WordDocument document) {
            ArgumentNullException.ThrowIfNull(document);
            _document = document;
            document.CompatibilitySettings = this;
        }

        /// <summary>Gets or sets the document's Word compatibility mode.</summary>
        public WordCompatibilityMode CompatibilityMode {
            get {
                OfficeXmlNode setting = FindMode(_document.GetSettingsRoot(create: true).WChild("compat"));
                string? value = setting.WVal();
                return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int mode)
                    && Enum.IsDefined(typeof(WordCompatibilityMode), mode)
                    ? (WordCompatibilityMode)mode
                    : WordCompatibilityMode.None;
            }
            set {
                OfficeXmlNode compatibility = EnsureCompatibility();
                OfficeXmlNode setting = FindMode(compatibility);
                if (value == WordCompatibilityMode.None) {
                    if (setting.Exists) setting.Remove();
                } else {
                    if (setting.IsNone) {
                        setting = WordSettingsXml.AddChild(compatibility, "compatSetting", WordSettingsXml.CompatibilityOrder);
                        setting.SetAttribute(W, "name", CompatibilityModeName);
                    }
                    setting.SetAttribute(W, "uri", CompatibilityModeUri);
                    setting.SetAttribute(W, "val", ((int)value).ToString(CultureInfo.InvariantCulture));
                }
                _document._ownedModel?.MarkChanged();
            }
        }

        /// <summary>
        /// Gets or sets whether text columns fill sequentially instead of balancing
        /// at the end of a section. The default is <see langword="false"/>.
        /// </summary>
        /// <remarks>
        /// This property reads and preserves the stored compatibility option.
        /// Word 2013 layout mode ignores the option when rendering columns.
        /// </remarks>
        public bool DoNotBalanceTextColumns {
            get => ReadOption("noColumnBalance");
            set => SetOption("noColumnBalance", value ? true : null);
        }

        private static OfficeXmlNode FindMode(OfficeXmlNode compatibility) {
            foreach (OfficeXmlNode setting in compatibility.Children(W, "compatSetting")) {
                if (setting.WAttr("name") == CompatibilityModeName) return setting;
            }
            return default;
        }

        private bool ReadOption(string name) => _document.GetSettingsRoot().WChild("compat").WChild(name).IsOn();

        private void SetOption(string name, bool? enabled) {
            if (enabled.HasValue) {
                OfficeXmlNode compatibility = EnsureCompatibility();
                WordSettingsXml.EnsureChild(compatibility, name, WordSettingsXml.CompatibilityOrder)
                    .SetAttribute(W, "val", enabled.Value ? "true" : "false");
            } else {
                OfficeXmlNode setting = _document.GetSettingsRoot().WChild("compat").WChild(name);
                if (setting.IsNone) return;
                setting.Remove();
            }
            _document._ownedModel?.MarkChanged();
        }

        private OfficeXmlNode EnsureCompatibility() {
            OfficeXmlNode settings = _document.GetSettingsRoot(create: true);
            if (settings.IsNone) throw new InvalidOperationException("The document does not contain a main document part.");
            return WordSettingsXml.EnsureChild(settings, "compat", WordSettingsXml.SettingsOrder);
        }
    }
}
