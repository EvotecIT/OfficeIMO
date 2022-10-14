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

        public void RemoveProtection() {
            if (this.ProtectionType != null) {
                DocumentProtection documentProtection = this._document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.OfType<DocumentProtection>().FirstOrDefault();
                documentProtection.Remove();
            }
        }

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
        public string ProtectionPassword {
            set {
                Security.ProtectWordDoc(this._document._wordprocessingDocument, value);
            }
        }
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

        ////Open Word Setting File
        //DocumentSettingsPart settingsPart = xmlDOc.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().First();
        ////Update Fields
        //UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
        //updateFields.Val = new OnOffValue(true);

        //settingsPart.Settings.PrependChild<UpdateFieldsOnOpen>(updateFields);
        //settingsPart.Settings.Save();

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

        public int? DefaultFontSize {
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


        public string Language {
            get {
                //if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles != null) {
                //    if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults != null) {
                //        if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault != null) {
                //            if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle != null) {
                //                if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.Languages != null) {
                //                    return this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.Languages.Val;
                //                }
                //            }
                //        }
                //    }
                //}
                var runPropertiesBaseStyle = GetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.Languages != null) {
                        return runPropertiesBaseStyle.Languages.Val;
                    }
                }
                return null;
            }
            set {
                //if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles != null) {
                //    if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults == null) {
                //        this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults = new DocDefaults();
                //    }

                //    if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault == null) {
                //        this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault = new RunPropertiesDefault();
                //    }

                //    if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle == null) {
                //        this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle = new RunPropertiesBaseStyle();
                //    }

                //    if (this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.Languages == null) {
                //        this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.Languages = new Languages();
                //    }

                //    this._document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.Languages.Val = value;
                //}
                var runPropertiesBaseStyle = SetDefaultStyleProperties();
                if (runPropertiesBaseStyle != null) {
                    if (runPropertiesBaseStyle.Languages == null) {
                        runPropertiesBaseStyle.Languages = new Languages();
                    }
                    runPropertiesBaseStyle.Languages.Val = value;
                } else {
                    throw new Exception("Could not set language. Styles not found.");
                }
            }
        }
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
