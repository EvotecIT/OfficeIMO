using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Editable pagenumber configuration attached to a section.</summary>
    public sealed class WordPageNumberSettings {
        private readonly WordSection _section;
        internal WordPageNumberSettings(WordSection section) => _section = section;
        private W.PageNumberType? Value => _section._sectionProperties.GetFirstChild<W.PageNumberType>();

        private W.PageNumberType Ensure() {
            var value = Value;
            if (value == null) {
                value = new W.PageNumberType();
                _section._sectionProperties.AddChild(value, true);
            }
            return value;
        }

        /// <summary>Gets or sets the first page number, or null when numbering continues.</summary>
        public int? StartNumber {
            get => Value?.Start?.Value;
            set {
                Ensure().Start = value;
            }
        }

        /// <summary>Gets or sets the page-number format.</summary>
        public WordNumberFormat? NumberingFormat {
            get => Value?.Format?.Value.ToOfficeEnum();
            set {
                Ensure().Format = value?.ToOpenXml();
            }
        }

    }

    /// <summary>Editable footnote configuration attached to a section.</summary>
    public sealed class WordFootnoteSettings {
        private readonly WordSection _section;
        internal WordFootnoteSettings(WordSection section) => _section = section;
        private W.FootnoteProperties? Value => _section._sectionProperties.GetFirstChild<W.FootnoteProperties>();

        private W.FootnoteProperties Ensure() {
            var value = Value;
            if (value == null) {
                value = new W.FootnoteProperties();
                _section._sectionProperties.AddChild(value, true);
            }
            return value;
        }

        private void Set<T>(T child, bool present) where T : DocumentFormat.OpenXml.OpenXmlElement {
            if (!present) { Value?.RemoveAllChildren<T>(); return; }
            var value = Ensure();
            value.RemoveAllChildren<T>();
            value.AddChild(child, true);
        }

        /// <summary>Gets or sets the footnote placement.</summary>
        public WordFootnotePosition? Position {
            get => Value?.FootnotePosition?.Val?.Value.ToOfficeEnum();
            set {
                Set(new W.FootnotePosition { Val = value?.ToOpenXml() }, value.HasValue);
            }
        }

        /// <summary>Gets or sets the numbering format.</summary>
        public WordNumberFormat? NumberingFormat {
            get => Value?.NumberingFormat?.Val?.Value.ToOfficeEnum();
            set {
                Set(new W.NumberingFormat { Val = value?.ToOpenXml() }, value.HasValue);
            }
        }

        /// <summary>Gets or sets the numbering restart rule.</summary>
        public WordNoteNumberRestart? NumberingRestart {
            get => Value?.NumberingRestart?.Val?.Value.ToOfficeEnum();
            set {
                Set(new W.NumberingRestart { Val = value?.ToOpenXml() }, value.HasValue);
            }
        }

        /// <summary>Gets or sets the first note number.</summary>
        public int? StartNumber {
            get => Value?.NumberingStart?.Val?.Value;
            set {
                if (value < 0 || value > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
                Set(new W.NumberingStart { Val = value.HasValue ? (ushort?)value.Value : null }, value.HasValue);
            }
        }

    }

    /// <summary>Editable endnote configuration attached to a section.</summary>
    public sealed class WordEndnoteSettings {
        private readonly WordSection _section;
        internal WordEndnoteSettings(WordSection section) => _section = section;
        private W.EndnoteProperties? Value => _section._sectionProperties.GetFirstChild<W.EndnoteProperties>();

        private W.EndnoteProperties Ensure() {
            var value = Value;
            if (value == null) {
                value = new W.EndnoteProperties();
                _section._sectionProperties.AddChild(value, true);
            }
            return value;
        }

        private void Set<T>(T child, bool present) where T : DocumentFormat.OpenXml.OpenXmlElement {
            if (!present) { Value?.RemoveAllChildren<T>(); return; }
            var value = Ensure();
            value.RemoveAllChildren<T>();
            value.AddChild(child, true);
        }

        /// <summary>Gets or sets the document-wide endnote placement; null restores document-end.</summary>
        public WordEndnotePosition? Position {
            get => _section._document.DocumentEndnotePosition.ToOfficeEnum();
            set {
                Set(new W.EndnotePosition { Val = value?.ToOpenXml() }, value.HasValue);
                _section._document.SetDocumentEndnotePosition(value?.ToOpenXml() ?? W.EndnotePositionValues.DocumentEnd);
            }
        }

        /// <summary>Gets or sets the numbering format.</summary>
        public WordNumberFormat? NumberingFormat {
            get => Value?.NumberingFormat?.Val?.Value.ToOfficeEnum();
            set {
                Set(new W.NumberingFormat { Val = value?.ToOpenXml() }, value.HasValue);
            }
        }

        /// <summary>Gets or sets the numbering restart rule.</summary>
        public WordNoteNumberRestart? NumberingRestart {
            get => Value?.NumberingRestart?.Val?.Value.ToOfficeEnum();
            set {
                Set(new W.NumberingRestart { Val = value?.ToOpenXml() }, value.HasValue);
            }
        }

        /// <summary>Gets or sets the first note number.</summary>
        public int? StartNumber {
            get => Value?.NumberingStart?.Val?.Value;
            set {
                if (value < 0 || value > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
                Set(new W.NumberingStart { Val = value.HasValue ? (ushort?)value.Value : null }, value.HasValue);
            }
        }

    }

    public partial class WordSection {
        /// <summary>Gets the editable configuration of the section's page-numbering configuration.</summary>
        public WordPageNumberSettings PageNumberSettings =>
            new WordPageNumberSettings(this);

        /// <summary>Gets the editable configuration of the section's footnote configuration.</summary>
        public WordFootnoteSettings FootnoteSettings =>
            new WordFootnoteSettings(this);

        /// <summary>Gets the editable configuration of section numbering and document-wide endnote placement.</summary>
        public WordEndnoteSettings EndnoteSettings =>
            new WordEndnoteSettings(this);
    }

    public partial class WordDocument {
        /// <summary>Gets the editable configuration of the first section's page-numbering configuration.</summary>
        public WordPageNumberSettings PageNumberSettings => Sections[0].PageNumberSettings;

        /// <summary>Gets the editable configuration of the first section's footnote configuration.</summary>
        public WordFootnoteSettings FootnoteSettings => Sections[0].FootnoteSettings;

        /// <summary>Gets the editable configuration of the first section's endnote configuration.</summary>
        public WordEndnoteSettings EndnoteSettings => Sections[0].EndnoteSettings;
    }
}
