using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    /// <summary>
    /// Predefined margin configurations available for a document section.
    /// </summary>
    public enum WordMargin {
        /// <summary>
        /// Standard one inch margins on all sides.
        /// </summary>
        Normal,
        /// <summary>
        /// Mirrored inside and outside margins for book binding.
        /// </summary>
        Mirrored,
        /// <summary>
        /// Slightly smaller left and right margins.
        /// </summary>
        Moderate,
        /// <summary>
        /// Minimal margins on all sides.
        /// </summary>
        Narrow,
        /// <summary>
        /// Extra wide margins on left and right.
        /// </summary>
        Wide,
        /// <summary>
        /// Margin preset used by Word 2003.
        /// </summary>
        Office2003Default,
        /// <summary>
        /// Custom margins not represented by other presets.
        /// </summary>
        Unknown
    }

    /// <summary>
    /// Provides access to page margin settings for a document section.
    /// </summary>
    public class WordMargins {
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private readonly WordSection _section;
        private static readonly string[] Attributes = { "top", "right", "bottom", "left", "header", "footer", "gutter" };
        private static readonly (WordMargin Type, int[] Twips)[] Presets = {
            (WordMargin.Normal, new[] { 1440, 1440, 1440, 1440, 720, 720, 0 }),
            (WordMargin.Mirrored, new[] { 1440, 1440, 1440, 1800, 720, 720, 0 }),
            (WordMargin.Moderate, new[] { 1440, 1080, 1440, 1080, 720, 720, 0 }),
            (WordMargin.Narrow, new[] { 720, 720, 720, 720, 720, 720, 0 }),
            (WordMargin.Wide, new[] { 1440, 2880, 1440, 2880, 720, 720, 0 }),
            (WordMargin.Office2003Default, new[] { 1440, 1800, 1440, 1800, 720, 720, 0 })
        };
        private static readonly int[] NormalTwips = Presets[0].Twips;

        /// <summary>Initializes a new instance for managing margins within the specified section.</summary>
        /// <param name="wordDocument">Parent document.</param>
        /// <param name="wordSection">Section associated with the margins.</param>
        public WordMargins(WordDocument wordDocument, WordSection wordSection) {
            _section = wordSection;
        }

        /// <summary>Gets or sets the left margin as a physical length.</summary>
        public OfficeLength? Left {
            get => GetLength("left", 1440);
            set => SetLength("left", value);
        }

        /// <summary>Gets or sets the right margin as a physical length.</summary>
        public OfficeLength? Right {
            get => GetLength("right", 1440);
            set => SetLength("right", value);
        }

        /// <summary>Gets or sets the top margin as a physical length.</summary>
        public OfficeLength? Top {
            get => GetLength("top", 1440, signed: true);
            set => SetLength("top", value, signed: true);
        }

        /// <summary>Gets or sets the bottom margin as a physical length.</summary>
        public OfficeLength? Bottom {
            get => GetLength("bottom", 1440, signed: true);
            set => SetLength("bottom", value, signed: true);
        }

        /// <summary>Gets or sets the header distance from the page edge.</summary>
        public OfficeLength? HeaderDistance {
            get => GetLength("header", 720);
            set => SetLength("header", value);
        }

        /// <summary>Gets or sets the footer distance from the page edge.</summary>
        public OfficeLength? FooterDistance {
            get => GetLength("footer", 720);
            set => SetLength("footer", value);
        }

        /// <summary>Gets or sets the binding gutter.</summary>
        public OfficeLength? Gutter {
            get => GetLength("gutter", 0);
            set => SetLength("gutter", value);
        }

        private OfficeLength GetLength(string name, int normal, bool signed = false) {
            string? token = _section.SectionNode.Child(W, "pgMar").Attribute(W, name);
            long? twips = signed ? (long?)OfficeXmlViewValues.ParseInt32(token) : OfficeXmlViewValues.ParseUInt32(token);
            return OfficeLength.Points((twips ?? normal) / 20D);
        }

        private void SetLength(string name, OfficeLength? value, bool signed = false) {
            if (!signed && value.HasValue && value.Value.ToEmus() < 0)
                throw new ArgumentOutOfRangeException(nameof(value), "The length cannot be negative.");
            long? twips = value?.ToTwips();
            string? token = signed
                ? OfficeXmlViewValues.Format(checked((int?)twips))
                : OfficeXmlViewValues.Format(checked((uint?)twips));
            OfficeXmlNode section = _section.EditSectionNode();
            OfficeXmlNode margin = section.Child(W, "pgMar");
            if (margin.IsNone) margin = AppendMargins(section, NormalTwips);
            margin.SetAttribute(W, name, token);
        }

        /// <summary>Gets or sets a predefined margin set that matches the current settings.</summary>
        public WordMargin? Type {
            get {
                OfficeXmlNode margins = _section.SectionNode.Child(W, "pgMar");
                if (margins.IsNone) return WordMargin.Normal;
                var twips = (Read(0), Read(1), Read(2), Read(3), Read(4), Read(5), Read(6));
                foreach (var preset in Presets) {
                    int[] expected = preset.Twips;
                    if (twips == (expected[0], expected[1], expected[2], expected[3], expected[4], expected[5], expected[6]))
                        return preset.Type;
                }
                return WordMargin.Unknown;

                int? Read(int index) => OfficeXmlViewValues.ParseInt32(margins.Attribute(W, Attributes[index]));
            }
            set {
                int[]? preset = null;
                if (value.HasValue) {
                    foreach (var candidate in Presets) {
                        if (candidate.Type == value.Value) {
                            preset = candidate.Twips;
                            break;
                        }
                    }
                    if (preset == null) throw new ArgumentOutOfRangeException(nameof(value));
                }
                OfficeXmlNode section = _section.EditSectionNode();
                OfficeXmlNode previous = section.Child(W, "pgMar");
                if (previous.Exists) previous.Remove();
                if (preset != null) AppendMargins(section, preset);
            }
        }

        private static OfficeXmlNode AppendMargins(OfficeXmlNode section, int[] twips) {
            OfficeXmlNode margin = section.CreateElement(W, "pgMar");
            for (int index = 0; index < Attributes.Length; index++) {
                margin.SetAttribute(W, Attributes[index], OfficeXmlViewValues.Format(twips[index]));
            }
            return section.Append(margin);
        }
    }
}
