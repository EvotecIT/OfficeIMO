using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides typed access to document-level settings such as protection,
    /// fonts and view options.
    /// </summary>
    public class WordSettings {
        private readonly WordDocument _document;

        /// <summary>
        /// Remove protection from document (if it's set).
        /// </summary>
        public void RemoveProtection() {
            if (ProtectionType != null) {
                DocumentProtection? documentProtection = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings?
                    .OfType<DocumentProtection>().FirstOrDefault();
                documentProtection?.Remove();
            }
        }

        /// <summary>
        /// Get or set Protection Type for the document
        /// </summary>
        public DocumentProtectionValues? ProtectionType {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings != null) {
                    DocumentProtection? documentProtection = settings
                        .OfType<DocumentProtection>()
                        .FirstOrDefault();
                    if (documentProtection != null) {
                        return documentProtection.Edit?.Value;
                    }
                }

                return null;
            }
            set {
                var settingsPart = _document._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart;
                if (settingsPart == null) {
                    return;
                }
                if (settingsPart.Settings == null) {
                    settingsPart.Settings = new Settings();
                }
                DocumentProtection? documentProtection = settingsPart.Settings
                    .OfType<DocumentProtection>()
                    .FirstOrDefault();
                if (documentProtection != null) {
                    documentProtection.Edit = value;
                } else {
                    throw new InvalidOperationException("Please first set password using 'ProtectionPassword' property before setting up encryption type.");
                }
            }
        }

        /// <summary>
        /// Set a Protection Password for the document
        /// </summary>
        public string ProtectionPassword {
            set {
                Security.ProtectWordDoc(_document._wordprocessingDocument, value);
            }
        }

        /// <summary>
        /// Get or set Zoom Preset for the document
        /// </summary>
        public PresetZoomValues? ZoomPreset {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings?.Zoom?.Val == null) {
                    return null;
                }
                return settings.Zoom.Val;

            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                if (settings.Zoom == null) {
                    settings.Zoom = new Zoom();
                }
                settings.Zoom.Val = value;
            }
        }

        /// <summary>
        /// Get or set Character Spacing Control
        /// </summary>
        public CharacterSpacingValues? CharacterSpacingControl {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                var characterSpacingControl = settings?
                    .OfType<CharacterSpacingControl>()
                    .FirstOrDefault();
                if (characterSpacingControl == null) {
                    return null;
                }

                return characterSpacingControl.Val?.Value;

            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                var characterSpacingControl = settings
                    .OfType<CharacterSpacingControl>()
                    .FirstOrDefault();
                if (characterSpacingControl == null) {
                    characterSpacingControl = new CharacterSpacingControl();
                    settings.Append(characterSpacingControl);
                }
                characterSpacingControl.Val = value;
            }
        }


        /// <summary>
        /// Get or set Default Tab Stop for the document
        /// </summary>
        public int DefaultTabStop {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                var defaultStop = settings?
                    .OfType<DefaultTabStop>()
                    .FirstOrDefault();
                if (defaultStop?.Val == null) {
                    return 0;
                }
                return defaultStop.Val;

            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                var defaultStop = settings
                    .OfType<DefaultTabStop>()
                    .FirstOrDefault();
                if (defaultStop == null) {
                    defaultStop = new DefaultTabStop();
                    settings.Append(defaultStop);
                }
                defaultStop.Val = (Int16Value)value;
            }
        }

        /// <summary>
        /// Get or Set Zoome Percentage for the document
        /// </summary>
        public int? ZoomPercentage {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                var percent = settings?.Zoom?.Percent;
                if (percent == null) {
                    return null;
                }
                return int.Parse(percent!, CultureInfo.InvariantCulture);
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                if (settings.Zoom == null) {
                    settings.Zoom = new Zoom();
                }
                settings.Zoom.Percent = value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : null;
            }
        }

        /// <summary>
        /// Tell Word to update fields when opening word.
        /// Without this option the document fields won't be refreshed until manual intervention.
        /// </summary>
        public bool UpdateFieldsOnOpen {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                var updateFieldsOnOpen = settings?
                    .GetFirstChild<UpdateFieldsOnOpen>();
                if (updateFieldsOnOpen == null) {
                    return false;
                }
                return updateFieldsOnOpen.Val?.Value ?? false;
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                var updateFieldsOnOpen = settings.GetFirstChild<UpdateFieldsOnOpen>();
                if (updateFieldsOnOpen == null) {
                    updateFieldsOnOpen = new UpdateFieldsOnOpen();
                    settings.PrependChild(updateFieldsOnOpen);
                }
                updateFieldsOnOpen.Val = value;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordSettings"/> class for the specified document.
        /// </summary>
        /// <param name="document">Document whose settings are managed.</param>
        public WordSettings(WordDocument document) {
            _ = document ?? throw new ArgumentNullException(nameof(document));
            if (document.FileOpenAccess != FileAccess.Read) {
                var mainPart = document._wordprocessingDocument.MainDocumentPart;
                if (mainPart == null) {
                    throw new InvalidOperationException("MainDocumentPart is missing.");
                }
                var documentSettingsPart = mainPart.DocumentSettingsPart;
                if (documentSettingsPart == null) {
                    documentSettingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                }

                var settings = documentSettingsPart.Settings;
                if (settings == null) {
                    settings = new Settings();
                    settings.Save(documentSettingsPart);
                }
            }
            _document = document;
            document.Settings = this;
        }

        private RunPropertiesBaseStyle? GetDefaultStyleProperties() {
            return _document._wordprocessingDocument.MainDocumentPart?
                .StyleDefinitionsPart?.Styles?
                .DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        }

        private RunPropertiesBaseStyle? SetDefaultStyleProperties() {
            var styles = _document._wordprocessingDocument.MainDocumentPart?
                .StyleDefinitionsPart?.Styles;
            if (styles == null) {
                return null;
            }

            if (styles.DocDefaults == null) {
                styles.DocDefaults = new DocDefaults();
            }

            if (styles.DocDefaults.RunPropertiesDefault == null) {
                styles.DocDefaults.RunPropertiesDefault = new RunPropertiesDefault();
            }

            if (styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle == null) {
                styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle = new RunPropertiesBaseStyle();
            }

            return styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle;
        }

        /// <summary>
        /// Gets or Sets default font size for the whole document. Default is 11.
        /// </summary>
        public int? FontSize {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    var fontSize = runPropertiesBaseStyle.FontSize?.Val;
                    if (fontSize != null) {
                        return int.Parse(fontSize!, CultureInfo.InvariantCulture) / 2;
                    }
                }
                return null;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.FontSize == null) {
                        runPropertiesBaseStyle.FontSize = new FontSize();
                    }
                    runPropertiesBaseStyle.FontSize.Val = value.HasValue ? (value.Value * 2).ToString(CultureInfo.InvariantCulture) : null;
                } else {
                    throw new InvalidOperationException("Could not set font size. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default font size complex script for the whole document. Default is 11.
        /// </summary>
        public int? FontSizeComplexScript {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    var fontSize = runPropertiesBaseStyle.FontSizeComplexScript?.Val;
                    if (fontSize != null) {
                        return int.Parse(fontSize!, CultureInfo.InvariantCulture) / 2;
                    }
                }
                return null;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.FontSizeComplexScript == null) {
                        runPropertiesBaseStyle.FontSizeComplexScript = new FontSizeComplexScript();
                    }
                    runPropertiesBaseStyle.FontSizeComplexScript.Val = value.HasValue ? (value.Value * 2).ToString(CultureInfo.InvariantCulture) : null;
                } else {
                    throw new InvalidOperationException("Could not set font size complex script. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default font family for the whole document.
        /// </summary>
        /// <seealso href="http://officeopenxml.com/WPtextFonts.php">WordProcessingText Fonts </seealso>
        public string? FontFamily {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                return runPropertiesBaseStyle?.RunFonts?.Ascii;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.RunFonts == null) {
                        runPropertiesBaseStyle.RunFonts = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                    }
                    runPropertiesBaseStyle.RunFonts.AsciiTheme = null;
                    runPropertiesBaseStyle.RunFonts.Ascii = value;
                    runPropertiesBaseStyle.RunFonts.HighAnsi = value;
                    runPropertiesBaseStyle.RunFonts.HighAnsiTheme = null;
                    runPropertiesBaseStyle.RunFonts.EastAsia = value;
                    runPropertiesBaseStyle.RunFonts.EastAsiaTheme = null;
                    runPropertiesBaseStyle.RunFonts.ComplexScript = value;
                    runPropertiesBaseStyle.RunFonts.ComplexScriptTheme = null;
                } else {
                    throw new InvalidOperationException("Could not set font family. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default font family for the whole document in HighAnsi.
        /// </summary>
        /// <seealso href="http://officeopenxml.com/WPtextFonts.php">WordProcessingText Fonts </seealso>
        public string? FontFamilyHighAnsi {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                return runPropertiesBaseStyle?.RunFonts?.HighAnsi;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.RunFonts == null) {
                        runPropertiesBaseStyle.RunFonts = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                    }
                    if (string.IsNullOrEmpty(value)) {
                        runPropertiesBaseStyle.RunFonts.HighAnsi = null;
                    } else {
                        runPropertiesBaseStyle.RunFonts.HighAnsi = value;
                    }
                    runPropertiesBaseStyle.RunFonts.HighAnsiTheme = null;
                } else {
                    throw new InvalidOperationException("Could not set font family. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default font family for the whole document in EastAsia.
        /// </summary>
        /// <seealso href="http://officeopenxml.com/WPtextFonts.php">WordProcessingText Fonts </seealso>
        public string? FontFamilyEastAsia {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                return runPropertiesBaseStyle?.RunFonts?.EastAsia;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.RunFonts == null) {
                        runPropertiesBaseStyle.RunFonts = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                    }
                    if (string.IsNullOrEmpty(value)) {
                        runPropertiesBaseStyle.RunFonts.EastAsia = null;
                    } else {
                        runPropertiesBaseStyle.RunFonts.EastAsia = value;
                    }
                    runPropertiesBaseStyle.RunFonts.EastAsiaTheme = null;
                } else {
                    throw new InvalidOperationException("Could not set font family. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default font family for the whole document in ComplexScript.
        /// </summary>
        /// <seealso href="http://officeopenxml.com/WPtextFonts.php">WordProcessingText Fonts </seealso>
        public string? FontFamilyComplexScript {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                return runPropertiesBaseStyle?.RunFonts?.ComplexScript;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.RunFonts == null) {
                        runPropertiesBaseStyle.RunFonts = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                    }
                    if (string.IsNullOrEmpty(value)) {
                        runPropertiesBaseStyle.RunFonts.ComplexScript = null;
                    } else {
                        runPropertiesBaseStyle.RunFonts.ComplexScript = value;
                    }
                    runPropertiesBaseStyle.RunFonts.ComplexScriptTheme = null;
                } else {
                    throw new InvalidOperationException("Could not set font family. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default language for the whole document. Default is en-Us.
        /// </summary>
        public string? Language {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                return runPropertiesBaseStyle?.Languages?.Val;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.Languages == null) {
                        runPropertiesBaseStyle.Languages = new Languages();
                    }
                    runPropertiesBaseStyle.Languages.Val = value;
                } else {
                    throw new InvalidOperationException("Could not set language. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default Background Color for the whole document
        /// </summary>
        public string? BackgroundColor {
            get {
                var background = _document._wordprocessingDocument.MainDocumentPart?
                    .Document?.DocumentBackground;
                return background?.Color?.Value?.ToLowerInvariant();
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                settings.DisplayBackgroundShape = new DisplayBackgroundShape();
                var document = _document._wordprocessingDocument.MainDocumentPart?.Document;
                if (document == null) {
                    return;
                }
                if (document.DocumentBackground == null) {
                    document.DocumentBackground = new DocumentBackground();
                }
                document.DocumentBackground.Color = value;
            }
        }
        /// <summary>
        /// Sets the background color using a hex value.
        /// </summary>
        /// <param name="backgroundColor">Hexadecimal color value.</param>
        /// <returns>The current <see cref="WordSettings"/> instance.</returns>
        public WordSettings SetBackgroundColor(string backgroundColor) {
            BackgroundColor = backgroundColor;
            return this;
        }

        /// <summary>
        /// Sets the background color using a <see cref="SixLabors.ImageSharp.Color"/> value.
        /// </summary>
        /// <param name="backgroundColor">Color value.</param>
        /// <returns>The current <see cref="WordSettings"/> instance.</returns>
        public WordSettings SetBackgroundColor(SixLabors.ImageSharp.Color backgroundColor) {
            BackgroundColor = backgroundColor.ToHexColor();
            return this;
        }

        /// <summary>
        /// Sets property in document recommending user to open the document as read only
        /// User can choose to do so, or ignore this recommendation
        /// </summary>
        public bool? ReadOnlyRecommended {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings?.WriteProtection?.Recommended == null) {
                    return false;
                }
                var recommended = settings.WriteProtection.Recommended;
                if (recommended.Value == true && (recommended.InnerText == "1" || string.IsNullOrEmpty(recommended.InnerText))) {
                    return true;
                }
                return false;
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                if (settings.WriteProtection == null && value != null && value != false) {
                    settings.WriteProtection = new WriteProtection();
                }
                if (settings.WriteProtection != null) {
                    if (value == null || value == false) {
                        settings.WriteProtection.Recommended = null;
                        if (string.IsNullOrEmpty(settings.WriteProtection.Hash) && settings.WriteProtection.Recommended == null) {
                            settings.WriteProtection.Remove();
                        }
                    } else {
                        var onOff = new DocumentFormat.OpenXml.OnOffValue(true) { InnerText = "1" };
                        settings.WriteProtection.Recommended = onOff;
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the document is marked as
        /// final. When set to <c>true</c>, Word notifies users that the
        /// document is finalized and should be treated as read-only.
        /// </summary>
        public bool FinalDocument {
            get {
                if (_document.CustomDocumentProperties.TryGetValue("_MarkAsFinal", out var markFinalProperty)) {
                    if (markFinalProperty?.Value is string valueString) {
                        return string.Equals(valueString, "1", StringComparison.Ordinal);
                    }
                }
                return false;
            }
            set {
                string newValue = value ? "1" : "0";
                if (_document.CustomDocumentProperties.ContainsKey("_MarkAsFinal")) {
                    _document.CustomDocumentProperties["_MarkAsFinal"].Value = newValue;
                } else {
                    _document.CustomDocumentProperties.Add("_MarkAsFinal", new WordCustomProperty(newValue));
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the gutter should be placed
        /// at the top of the page when the document uses a vertical layout.
        /// </summary>
        public bool GutterAtTop {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                var gutterAtTop = settings?.GetFirstChild<GutterAtTop>();
                if (gutterAtTop == null) {
                    return false;
                }
                return gutterAtTop.Val?.Value ?? false;
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                var gutterAtTop = settings.GetFirstChild<GutterAtTop>();
                if (gutterAtTop == null) {
                    gutterAtTop = new GutterAtTop();
                    settings.Append(gutterAtTop);
                }
                gutterAtTop.Val = value;
            }
        }

        /// <summary>
        /// Enable or disable tracking of revisions in the document.
        /// </summary>
        public bool TrackRevisions {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                return settings?.GetFirstChild<TrackRevisions>() != null;
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                var track = settings.GetFirstChild<TrackRevisions>();
                if (value) {
                    if (track == null) {
                        settings.Append(new TrackRevisions());
                    }
                } else {
                    track?.Remove();
                }
            }
        }

        /// <summary>
        /// Enable or disable tracking of comments in the document.
        /// Wrapper around <see cref="TrackRevisions"/> for backwards compatibility.
        /// </summary>
        public bool TrackComments {
            get => TrackRevisions;
            set => TrackRevisions = value;
        }

        /// <summary>
        /// Enable or disable tracking of formatting changes.
        /// </summary>
        public bool TrackFormatting {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                return settings?.GetFirstChild<DoNotTrackFormatting>() == null;
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                var formatting = settings.GetFirstChild<DoNotTrackFormatting>();
                if (value) {
                    formatting?.Remove();
                } else if (formatting == null) {
                    settings.Append(new DoNotTrackFormatting());
                }
            }
        }

        /// <summary>
        /// Enable or disable tracking of move operations.
        /// </summary>
        public bool TrackMoves {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                return settings?.GetFirstChild<DoNotTrackMoves>() == null;
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?
                    .DocumentSettingsPart?.Settings;
                if (settings == null) {
                    return;
                }
                var moves = settings.GetFirstChild<DoNotTrackMoves>();
                if (value) {
                    moves?.Remove();
                } else if (moves == null) {
                    settings.Append(new DoNotTrackMoves());
                }
            }
        }
    }
}
