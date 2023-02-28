using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordSettings {
        private WordDocument _document;

        /// <summary>
        /// Remove protection from document (if it's set).
        /// </summary>
        public void RemoveProtection() {
            if (this.ProtectionType != null) {
                DocumentProtection documentProtection = this._document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.OfType<DocumentProtection>().FirstOrDefault();
                documentProtection.Remove();
            }
        }

        /// <summary>
        /// Get or set Protection Type for the document
        /// </summary>
        public DocumentProtectionValues? ProtectionType {
            get {
                if (this._document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings != null) {
                    DocumentProtection documentProtection = this._document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.OfType<DocumentProtection>().FirstOrDefault();
                    if (documentProtection != null) {
                        return documentProtection.Edit;
                    }
                }

                return null;
            }
            set {
                if (this._document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings == null) {
                    this._document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings = new Settings();
                }
                DocumentProtection documentProtection = this._document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.OfType<DocumentProtection>().FirstOrDefault();
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
                Security.ProtectWordDoc(this._document._wordprocessingDocument, value);
            }
        }

        /// <summary>
        /// Get or set Zoom Preset for the document
        /// </summary>
        public PresetZoomValues? ZoomPreset {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                if (settings.Zoom == null) {
                    return null;
                }

                if (settings.Zoom.Val == null) {
                    return null;
                }
                return settings.Zoom.Val;

            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
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
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                var characterSpacingControl = settings.OfType<CharacterSpacingControl>().FirstOrDefault();
                if (characterSpacingControl == null) {
                    return null;
                }

                return characterSpacingControl.Val;

            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                var characterSpacingControl = settings.OfType<CharacterSpacingControl>().FirstOrDefault();
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
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                var defaultStop = settings.OfType<DefaultTabStop>().FirstOrDefault();
                if (defaultStop == null) {
                    return 0;
                }
                if (defaultStop.Val == null) {
                    return 0;
                }
                return defaultStop.Val;

            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                var defaultStop = settings.OfType<DefaultTabStop>().FirstOrDefault();
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
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                if (settings.Zoom == null) {
                    return null;
                }
                if (settings.Zoom.Percent == null) {
                    return null;
                }
                return int.Parse(settings.Zoom.Percent);
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                if (settings.Zoom == null) {
                    settings.Zoom = new Zoom();
                }
                settings.Zoom.Percent = value.ToString();
            }
        }

        /// <summary>
        /// Tell Word to update fields when opening word.
        /// Without this option the document fields won't be refreshed until manual intervention.
        /// </summary>
        public bool UpdateFieldsOnOpen {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                var updateFieldsOnOpen = settings.GetFirstChild<UpdateFieldsOnOpen>();
                if (updateFieldsOnOpen == null) {
                    return false;
                }
                return updateFieldsOnOpen.Val;
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                var updateFieldsOnOpen = settings.GetFirstChild<UpdateFieldsOnOpen>();
                if (updateFieldsOnOpen == null) {
                    updateFieldsOnOpen = new UpdateFieldsOnOpen();
                    settings.PrependChild<UpdateFieldsOnOpen>(updateFieldsOnOpen);
                }
                updateFieldsOnOpen.Val = value;
            }
        }

        public WordSettings(WordDocument document) {
            if (document.FileOpenAccess != FileAccess.Read) {
                var documentSettingsPart = document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart;
                if (documentSettingsPart == null) {
                    documentSettingsPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                }

                var settings = document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
                if (settings == null) {
                    settings = new Settings();
                    settings.Save(documentSettingsPart);
                }
            }
            _document = document;
            document.Settings = this;
        }

        private RunPropertiesBaseStyle? GetDefaultStyleProperties() {
            if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles != null) {
                if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults != null) {
                    if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault != null) {
                        if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle != null) {
                            if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle != null) {
                                return this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle;
                            }
                        }
                    }
                }
            }
            return null;
        }

        private RunPropertiesBaseStyle? SetDefaultStyleProperties() {
            if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles != null) {
                if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults == null) {
                    this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults = new DocDefaults();
                }

                if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault == null) {
                    this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault = new RunPropertiesDefault();
                }

                if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle == null) {
                    this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle = new RunPropertiesBaseStyle();
                }

                return this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle;
            }

            return null;
        }

        /// <summary>
        /// Gets or Sets default font size for the whole document. Default is 11.
        /// </summary>
        public int? FontSize {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.FontSize != null) {
                        var fontSize = runPropertiesBaseStyle.FontSize.Val;
                        return int.Parse(fontSize) / 2;
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
                    runPropertiesBaseStyle.FontSize.Val = (value * 2).ToString();
                } else {
                    throw new Exception("Could not set font size. Styles not found.");
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
                    if (runPropertiesBaseStyle.FontSizeComplexScript != null) {
                        var fontSize = runPropertiesBaseStyle.FontSizeComplexScript.Val;
                        return int.Parse(fontSize) / 2;
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
                    runPropertiesBaseStyle.FontSizeComplexScript.Val = (value * 2).ToString();
                } else {
                    throw new Exception("Could not set font size complex script. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default font family for the whole document.
        /// </summary>
        /// <seealso href="http://officeopenxml.com/WPtextFonts.php">WordProcessingText Fonts </seealso>
        public string FontFamily {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.RunFonts != null) {
                        var fontFamily = runPropertiesBaseStyle.RunFonts.Ascii;
                        return fontFamily;
                    }
                }
                return null;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    //runPropertiesBaseStyle.RunFonts = new RunFonts();
                    if (runPropertiesBaseStyle.RunFonts == null) {
                        runPropertiesBaseStyle.RunFonts = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                    }
                    // we need to reset default AsciiTheme, before applying Ascii
                    runPropertiesBaseStyle.RunFonts.AsciiTheme = null;
                    runPropertiesBaseStyle.RunFonts.Ascii = value;
                } else {
                    throw new Exception("Could not set font family. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default font family for the whole document in highAnsi.
        /// </summary>
        /// <seealso href="http://officeopenxml.com/WPtextFonts.php">WordProcessingText Fonts </seealso>
        public string FontFamilyHighAnsi {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.RunFonts != null) {
                        var fontFamily = runPropertiesBaseStyle.RunFonts.HighAnsi;
                        return fontFamily;
                    }
                }
                return null;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    //runPropertiesBaseStyle.RunFonts = new RunFonts();
                    if (runPropertiesBaseStyle.RunFonts == null) {
                        runPropertiesBaseStyle.RunFonts = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                    }
                    // we also need to change it in highAnsi to fix https://github.com/EvotecIT/OfficeIMO/issues/54
                    runPropertiesBaseStyle.RunFonts.HighAnsi = value;
                    runPropertiesBaseStyle.RunFonts.HighAnsiTheme = null;
                } else {
                    throw new Exception("Could not set font family. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default language for the whole document. Default is en-Us.
        /// </summary>
        public string Language {
            get {
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.Languages != null) {
                        return runPropertiesBaseStyle.Languages.Val;
                    }
                }
                return null;
            }
            set {
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.Languages == null) {
                        runPropertiesBaseStyle.Languages = new Languages();
                    }
                    runPropertiesBaseStyle.Languages.Val = value;
                    //runPropertiesBaseStyle.Languages.EastAsia = value;
                } else {
                    throw new Exception("Could not set language. Styles not found.");
                }
            }
        }

        /// <summary>
        /// Gets or Sets default Background Color for the whole document
        /// </summary>
        public string BackgroundColor {
            get {
                if (_document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground != null) {
                    return _document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground.Color;
                }

                return null;
            }
            set {
                _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.DisplayBackgroundShape = new DisplayBackgroundShape();
                if (_document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground == null) {
                    _document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground = new DocumentBackground();
                }
                _document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground.Color = value;
            }
        }
        public WordSettings SetBackgroundColor(string backgroundColor) {
            this.BackgroundColor = backgroundColor;
            return this;
        }
        public WordSettings SetBackgroundColor(SixLabors.ImageSharp.Color backgroundColor) {
            this.BackgroundColor = backgroundColor.ToHexColor();
            return this;
        }

    }
}
