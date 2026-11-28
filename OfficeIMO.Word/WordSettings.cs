using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides typed access to document-level settings such as protection,
    /// fonts and view options.
    /// </summary>
    public partial class WordSettings {
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private readonly WordDocument _document;

        /// <summary>Initializes a settings manager for the specified document.</summary>
        /// <param name="document">Document whose settings are managed.</param>
        public WordSettings(WordDocument document) {
            ArgumentNullException.ThrowIfNull(document);
            if (document.FileOpenAccess != FileAccess.Read) _ = document.GetSettingsRoot(create: true);
            _document = document;
            document.Settings = this;
        }

        /// <summary>Removes the document's editing protection.</summary>
        public void RemoveProtection() {
            if (ProtectionType.HasValue) RemoveSetting("documentProtection");
        }

        /// <summary>Gets or sets the editing restrictions for the document.</summary>
        public WordDocumentProtectionType? ProtectionType {
            get => ReadSetting("documentProtection").WAttr("edit") switch {
                "none" => WordDocumentProtectionType.None,
                "readOnly" => WordDocumentProtectionType.ReadOnly,
                "comments" => WordDocumentProtectionType.Comments,
                "trackedChanges" => WordDocumentProtectionType.TrackedChanges,
                "forms" => WordDocumentProtectionType.Forms,
                null => null,
                _ => throw InvalidEnumeration()
            };
            set {
                OfficeXmlNode settings = _document.GetSettingsRoot();
                if (settings.IsNone) return;
                OfficeXmlNode protection = settings.WChild("documentProtection");
                if (protection.IsNone) {
                    throw new InvalidOperationException("Please first set password using 'ProtectionPassword' property before setting up encryption type.");
                }
                string? token = value switch {
                    WordDocumentProtectionType.None => "none",
                    WordDocumentProtectionType.ReadOnly => "readOnly",
                    WordDocumentProtectionType.Comments => "comments",
                    WordDocumentProtectionType.TrackedChanges => "trackedChanges",
                    WordDocumentProtectionType.Forms => "forms",
                    null => null,
                    _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown document protection type.")
                };
                protection.SetAttribute(W, "edit", token);
                MarkChanged();
            }
        }

        /// <summary>Sets document protection using an OfficeIMO-owned value.</summary>
        public WordSettings SetProtectionType(WordDocumentProtectionType protectionType) {
            ProtectionType = protectionType;
            return this;
        }

        /// <summary>Sets the password for the document's editing restrictions.</summary>
        public string ProtectionPassword {
            set {
                OfficeXmlNode settings = _document.GetSettingsRoot();
                if (settings.IsNone) return;
                if (_document.FileOpenAccess == FileAccess.Read) throw new IOException("Cannot change document protection on a read-only document.");
                OfficeXmlNode protection = settings.CreateElement(W, "documentProtection");
                Security.ProtectWordDoc(protection, value);
                WordSettingsXml.SetChild(settings, protection, WordSettingsXml.SettingsOrder);
                MarkChanged();
            }
        }

        /// <summary>Gets or sets the zoom preset.</summary>
        public WordZoomPreset? ZoomPreset {
            get => ReadSetting("zoom").WVal() is { } token
                ? WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordZoomPreset value) ? value : throw InvalidEnumeration()
                : null;
            set => SetSettingAttribute("zoom", "val", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets character spacing control.</summary>
        public WordCharacterSpacing? CharacterSpacingControl {
            get => ReadSetting("characterSpacingControl").WVal() is { } token
                ? WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordCharacterSpacing value) ? value : throw InvalidEnumeration()
                : null;
            set => SetSettingAttribute("characterSpacingControl", "val", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the default tab stop, in twentieths of a point.</summary>
        public int DefaultTabStop {
            get => ReadSetting("defaultTabStop").WVal() is { } value ? short.Parse(value, CultureInfo.InvariantCulture) : 0;
            set => SetSettingAttribute("defaultTabStop", "val", unchecked((short)value).ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>Gets or sets the zoom percentage.</summary>
        public int? ZoomPercentage {
            get => ReadSetting("zoom").WAttr("percent") is { } percent ? int.Parse(percent, CultureInfo.InvariantCulture) : null;
            set => SetSettingAttribute("zoom", "percent", value?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Tells Word to update fields when opening the document.
        /// Without this option the document fields require a manual refresh.
        /// </summary>
        public bool UpdateFieldsOnOpen {
            get => ReadSetting("updateFields").WVal() is "true" or "1" or "on";
            set {
                if (SetSettingAttribute("updateFields", "val", value ? "true" : "false")) {
                    _document.ReleaseTableOfContentUpdateFieldsOnOpen();
                }
            }
        }

        /// <summary>Gets or sets the default background color. Automatic color reads as <see langword="null"/>.</summary>
        public OfficeColor? BackgroundColor {
            get => Helpers.ParseNativeColor(_document.SettingsDocumentRoot.WChild("background").WAttr("color"));
            set {
                OfficeXmlNode settings = _document.GetSettingsRoot();
                if (settings.IsNone) return;
                WordSettingsXml.SetChild(settings, settings.CreateElement(W, "displayBackgroundShape"), WordSettingsXml.SettingsOrder);
                OfficeXmlNode document = _document.SettingsDocumentRoot;
                if (document.Exists) {
                    WordSettingsXml.EnsureChild(document, "background", WDocumentView.Order)
                        .SetAttribute(W, "color", value?.ToRgbHex());
                }
                MarkChanged();
            }
        }

        /// <summary>Sets the document's background color.</summary>
        /// <param name="backgroundColor">Color value.</param>
        /// <returns>The current <see cref="WordSettings"/> instance.</returns>
        public WordSettings SetBackgroundColor(OfficeColor backgroundColor) {
            BackgroundColor = backgroundColor;
            return this;
        }

        /// <summary>Gets or sets whether Word recommends opening the document as read-only.</summary>
        public bool? ReadOnlyRecommended {
            get {
                string? value = ReadSetting("writeProtection").WAttr("recommended");
                return WordXml.ParseOnOff(value) == true && value == "1";
            }
            set {
                if (value == true) {
                    SetSettingAttribute("writeProtection", "recommended", "1");
                } else {
                    OfficeXmlNode protection = ReadSetting("writeProtection");
                    if (protection.IsNone) return;
                    if (string.IsNullOrEmpty(protection.WAttr("hash"))) protection.Remove();
                    else protection.RemoveAttribute(W, "recommended");
                    MarkChanged();
                }
            }
        }

        /// <summary>
        /// Gets or sets whether the document is marked as final.
        /// Word notifies users that a finalized document should be treated as read-only.
        /// </summary>
        public bool FinalDocument {
            get => _document.CustomProperties.TryGetValue("_MarkAsFinal", out var property) && property?.Value is "1";
            set {
                string newValue = value ? "1" : "0";
                if (_document.CustomProperties.TryGetValue("_MarkAsFinal", out var property)) {
                    property.Value = newValue;
                } else {
                    _document.CustomProperties.Add("_MarkAsFinal", new WordCustomProperty(newValue));
                }
            }
        }

        /// <summary>Gets or sets whether Word tracks revisions.</summary>
        public bool TrackRevisions {
            get => ReadSetting("trackRevisions").Exists;
            set => SetSettingPresence("trackRevisions", value);
        }

        /// <summary>Gets or sets whether Word tracks formatting changes.</summary>
        public bool TrackFormatting {
            get => ReadSetting("doNotTrackFormatting").IsNone;
            set => SetSettingPresence("doNotTrackFormatting", !value);
        }

        /// <summary>Gets or sets whether Word tracks move operations.</summary>
        public bool TrackMoves {
            get => ReadSetting("doNotTrackMoves").IsNone;
            set => SetSettingPresence("doNotTrackMoves", !value);
        }

        private OfficeXmlNode ReadSetting(string name) => _document.GetSettingsRoot().WChild(name);

        private bool SetSettingAttribute(string name, string attributeName, string? value) {
            OfficeXmlNode settings = _document.GetSettingsRoot();
            if (settings.IsNone) return false;
            WordSettingsXml.EnsureChild(settings, name, WordSettingsXml.SettingsOrder).SetAttribute(W, attributeName, value);
            MarkChanged();
            return true;
        }

        private void SetSettingPresence(string name, bool present) {
            if (!present) {
                RemoveSetting(name);
                return;
            }
            OfficeXmlNode settings = _document.GetSettingsRoot();
            if (settings.IsNone || settings.WChild(name).Exists) return;
            WordSettingsXml.AddChild(settings, name, WordSettingsXml.SettingsOrder);
            MarkChanged();
        }

        private void RemoveSetting(string name) {
            OfficeXmlNode setting = ReadSetting(name);
            if (setting.IsNone) return;
            setting.Remove();
            MarkChanged();
        }

        private void MarkChanged() => _document._ownedModel?.MarkChanged();

        private static FormatException InvalidEnumeration() => new("The text value is not a valid enumeration value.");
    }
}
