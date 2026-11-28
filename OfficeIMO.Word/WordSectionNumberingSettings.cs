using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word {
    /// <summary>Editable pagenumber configuration attached to a section.</summary>
    public sealed class WordPageNumberSettings {
        private readonly WordSection _section;
        internal WordPageNumberSettings(WordSection section) => _section = section;
        private const int Ns = OfficeXmlNamespaces.WordprocessingML;
        private OfficeXmlNode Value => WordSection.FindSectionChild(_section.SectionNode, "pgNumType");
        private OfficeXmlNode Ensure() => _section.EnsureSectionChild("pgNumType");

        /// <summary>Gets or sets the first page number, or null when numbering continues.</summary>
        public int? StartNumber {
            get => OfficeXmlViewValues.ParseInt32(Value.Attribute(Ns, "start"));
            set => Ensure().SetAttribute(Ns, "start", OfficeXmlViewValues.Format(value));
        }

        /// <summary>Gets or sets the page-number format.</summary>
        public WordNumberFormat? NumberingFormat {
            get => Value.Attribute(Ns, "fmt") is string token &&
                WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordNumberFormat format)
                ? format : (WordNumberFormat?)null;
            set => Ensure().SetAttribute(Ns, "fmt", value?.ToOpenXmlToken());
        }
    }

    internal readonly struct SectionNoteSettings {
        private const int Ns = OfficeXmlNamespaces.WordprocessingML;
        internal static readonly OfficeXmlChildOrder Order = new(
            new[] { Ns, Ns, Ns, Ns }, new[] { "pos", "numFmt", "numStart", "numRestart" }, new[] { 0, 1, 2, 3 });

        private readonly WordSection _section;
        private readonly string _localName;

        internal SectionNoteSettings(WordSection section, string localName) {
            _section = section;
            _localName = localName;
        }

        internal OfficeXmlNode Edit() => _section.EnsureSectionChild(_localName);

        internal string? Get(string childName) => WordSection.FindSectionChild(_section.SectionNode, _localName).Child(Ns, childName).Attribute(Ns, "val");

        internal void Set(string name, string? token) {
            OfficeXmlNode node = token != null ? Edit() : WordSection.FindSectionChild(_section.EditSectionNode(), _localName);
            Set(node, name, token);
        }

        internal static void Set(OfficeXmlNode properties, string name, string? token) {
            if (properties.IsNone) return;
            properties.RemoveChildren(Ns, name);
            if (token != null) properties.AddOrdered(Ns, name, Order).SetAttribute(Ns, "val", token);
        }

        internal WordNumberFormat? NumberingFormat {
            get => Get("numFmt") is { } token &&
                WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordNumberFormat format)
                ? format : (WordNumberFormat?)null;
            set => Set("numFmt", value?.ToOpenXmlToken());
        }

        internal WordNoteNumberRestart? NumberingRestart {
            get => Get("numRestart") is { } token &&
                WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordNoteNumberRestart restart)
                ? restart : (WordNoteNumberRestart?)null;
            set => Set("numRestart", value?.ToOpenXmlToken());
        }

        internal int? StartNumber {
            get => Get("numStart") is { } token ? ushort.Parse(token, CultureInfo.InvariantCulture) : (int?)null;
            set => Set("numStart", StartToken(value));
        }

        internal static string? StartToken(int? value) {
            if (value < 0 || value > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
            return OfficeXmlViewValues.Format(value);
        }
    }

    /// <summary>Editable footnote configuration attached to a section.</summary>
    public sealed class WordFootnoteSettings {
        private SectionNoteSettings _settings;
        internal WordFootnoteSettings(WordSection section) => _settings = new SectionNoteSettings(section, "footnotePr");

        /// <summary>Gets or sets the footnote placement.</summary>
        public WordFootnotePosition? Position {
            get => _settings.Get("pos") is { } token &&
                WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordFootnotePosition position)
                ? position : (WordFootnotePosition?)null;
            set => _settings.Set("pos", value?.ToOpenXmlToken());
        }

        /// <summary>Gets or sets the numbering format.</summary>
        public WordNumberFormat? NumberingFormat { get => _settings.NumberingFormat; set => _settings.NumberingFormat = value; }

        /// <summary>Gets or sets the numbering restart rule.</summary>
        public WordNoteNumberRestart? NumberingRestart { get => _settings.NumberingRestart; set => _settings.NumberingRestart = value; }

        /// <summary>Gets or sets the first note number.</summary>
        public int? StartNumber { get => _settings.StartNumber; set => _settings.StartNumber = value; }
    }

    /// <summary>Editable endnote configuration attached to a section.</summary>
    public sealed class WordEndnoteSettings {
        private readonly WordSection _section;
        private SectionNoteSettings _settings;
        internal WordEndnoteSettings(WordSection section) {
            _section = section;
            _settings = new SectionNoteSettings(section, "endnotePr");
        }

        /// <summary>Gets or sets document-wide endnote placement. Null restores document-end placement.</summary>
        public WordEndnotePosition? Position {
            get => _section.ReadDocumentEndnotePlacement();
            set {
                _settings.Set("pos", value?.ToOpenXmlToken());
                _section.WriteDocumentEndnotePlacement(value ?? WordEndnotePosition.DocumentEnd);
            }
        }

        /// <summary>Gets or sets the numbering format.</summary>
        public WordNumberFormat? NumberingFormat { get => _settings.NumberingFormat; set => _settings.NumberingFormat = value; }

        /// <summary>Gets or sets the numbering restart rule.</summary>
        public WordNoteNumberRestart? NumberingRestart { get => _settings.NumberingRestart; set => _settings.NumberingRestart = value; }

        /// <summary>Gets or sets the first note number.</summary>
        public int? StartNumber { get => _settings.StartNumber; set => _settings.StartNumber = value; }
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
