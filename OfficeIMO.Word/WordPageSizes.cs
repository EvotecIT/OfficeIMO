using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Globalization;

namespace OfficeIMO.Word {
    /// <summary>
    /// PaperSizes used in Microsoft Office
    /// </summary>
    public enum WordPageSize {
        /// <summary>
        /// Custom/unknown paper size that is not defined within OfficeIMO
        /// </summary>
        Unknown,
        /// <summary>
        /// Letter is part of the North American loose paper size series.
        /// It is the standard for business and academic documents, and measures 216 × 279 mm or 8.5 × 11 inches.
        /// </summary>
        Letter,
        /// <summary>
        /// Legal is part of the North American loose paper size series.
        /// It is used to make legal pads, and measures 216 × 356 mm or 8.5 × 14 inches.
        /// </summary>
        Legal,
        /// <summary>
        /// Statement paper size 5.5 x 8.5 inches
        /// </summary>
        Statement,
        /// <summary>
        /// Executive paper size 7.25 x 10.50 inches.
        /// </summary>
        Executive,
        /// <summary>
        /// An A3 piece of paper measures 297 × 420 mm or 11.7 × 16.5 inches.
        /// </summary>
        A3,
        /// <summary>
        /// An A4 piece of paper measures 210 × 297 mm or 8.3 × 11.7 inches
        /// </summary>
        A4,
        /// <summary>
        /// An A5 piece of paper measures 148 × 210 mm or 5.8 × 8.3 inches.
        /// </summary>
        A5,
        /// <summary>
        /// An A6 piece of paper measures 105 × 148 mm or 4.1 × 5.8 inches.
        /// </summary>
        A6,
        /// <summary>
        /// JIS B5 paper, 182 × 257 millimeters. This retains the established B5 preset.
        /// </summary>
        B5,
        /// <summary>US Tabloid, 11 × 17 inches.</summary>
        Tabloid,
        /// <summary>JIS B4, 257 × 364 millimeters.</summary>
        B4Jis,
        /// <summary>Number 9 envelope, 3.875 × 8.875 inches.</summary>
        Envelope9,
        /// <summary>Number 10 envelope, 4.125 × 9.5 inches.</summary>
        Envelope10,
        /// <summary>US C sheet, 17 × 22 inches.</summary>
        CSheet,
        /// <summary>DL envelope, 110 × 220 millimeters.</summary>
        EnvelopeDl,
        /// <summary>C5 envelope, 162 × 229 millimeters.</summary>
        EnvelopeC5,
        /// <summary>C4 envelope, 229 × 324 millimeters.</summary>
        EnvelopeC4,
        /// <summary>B5 envelope, 176 × 250 millimeters.</summary>
        EnvelopeB5,
        /// <summary>Monarch envelope, 3.875 × 7.5 inches.</summary>
        EnvelopeMonarch
    }

    /// <summary>Dimensions and Word paper code for a built-in page-size preset.</summary>
    public sealed class WordPageSizeDefinition {
        internal WordPageSizeDefinition(uint widthTwips, uint heightTwips, ushort paperCode) {
            Width = global::OfficeIMO.OfficeLength.Points((widthTwips) / 20D);
            Height = global::OfficeIMO.OfficeLength.Points((heightTwips) / 20D);
            PaperCode = paperCode;
        }

        /// <summary>Gets the page width as a physical length.</summary>
        public global::OfficeIMO.OfficeLength Width { get; }

        /// <summary>Gets the page height as a physical length.</summary>
        public global::OfficeIMO.OfficeLength Height { get; }

        /// <summary>Gets the Word paper-size code.</summary>
        public ushort PaperCode { get; }
    }

    /// <summary>
    /// Provides helpers for manipulating Word page size and orientation.
    /// </summary>
    public partial class WordPageSizes {
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private static readonly WordPageSize[] Presets = global::OfficeIMO.Internal.EnumCompat.GetValues<WordPageSize>();
        private readonly WordSection _section;

        private OfficeXmlNode SizeNode => _section.SectionNode.Child(W, "pgSz");

        /// <summary>Gets or sets the preset paper size for the section.</summary>
        public WordPageSize? PageSize {
            get {
                OfficeXmlNode node = SizeNode;
                if (node.IsNone) return null;
                uint? width = ReadDimension(node, "w");
                uint? height = ReadDimension(node, "h");
                ushort? code = ReadCode(node);
                foreach (WordPageSize preset in Presets) {
                    if (GetDefinition(preset) is not { } definition) continue;
                    if (code.HasValue && definition.PaperCode != code.Value) continue;
                    long expectedWidth = definition.Width.ToTwips();
                    long expectedHeight = definition.Height.ToTwips();
                    // The paper code is optional. Accept a one-twip producer rounding difference.
                    if ((Matches(expectedWidth, width) && Matches(expectedHeight, height)) ||
                        (Matches(expectedWidth, height) && Matches(expectedHeight, width))) return preset;
                }
                return WordPageSize.Unknown;
            }
            set {
                WordPageSizeDefinition? preset = value.HasValue ? GetDefinition(value.Value) : null;
                OfficeXmlNode section = _section.EditSectionNode();
                OfficeXmlNode previous = section.Child(W, "pgSz");
                OfficePageOrientation orientation = ReadOrientation(previous);
                if (previous.Exists) previous.Remove();
                if (preset == null) return;
                OfficeXmlNode pageSize = WordSection.AddSectionChild(section, "pgSz");
                WritePreset(pageSize, preset);
                if (orientation == OfficePageOrientation.Landscape) ApplyOrientation(pageSize, orientation);
            }
        }

        private static bool Matches(long expected, uint? actual) => actual.HasValue && Math.Abs(expected - actual.Value) <= 1L;

        /// <summary>
        /// Applies a physical paper size in the requested orientation, or the section's current orientation when omitted.
        /// </summary>
        /// <remarks>Unlike individual dimension edits, both dimensions are oriented together. Clears the printer paper code.</remarks>
        internal void SetSize(OfficeIMO.Drawing.OfficePageSize size, OfficePageOrientation? orientation = null) {
            if (size.Width <= OfficeLength.Zero || size.Height <= OfficeLength.Zero)
                throw new ArgumentOutOfRangeException(nameof(size), "Page dimensions must be positive.");
            OfficePageOrientation effectiveOrientation = orientation ?? Orientation;
            string token = effectiveOrientation.ToOpenXmlToken();
            var orientedSize = effectiveOrientation == OfficePageOrientation.Landscape ? size.Landscape() : size.Portrait();
            uint width = checked((uint)orientedSize.Width.ToTwips());
            uint height = checked((uint)orientedSize.Height.ToTwips());
            OfficeXmlNode pageSize = EnsurePageSize();
            pageSize.SetAttribute(W, "w", OfficeXmlViewValues.Format(width));
            pageSize.SetAttribute(W, "h", OfficeXmlViewValues.Format(height));
            pageSize.SetAttribute(W, "orient", token);
            pageSize.SetAttribute(W, "code", null);
        }

        private OfficeXmlNode EnsurePageSize() => _section.EnsureSectionChild("pgSz");

        /// <summary>Gets or sets the section width. Explicit edits clear the paper code.</summary>
        public OfficeLength? Width {
            get => GetLength("w");
            set => SetLength("w", value);
        }

        /// <summary>Gets or sets the section height. Explicit edits clear the paper code.</summary>
        public OfficeLength? Height {
            get => GetLength("h");
            set => SetLength("h", value);
        }

        private OfficeLength? GetLength(string attribute) =>
            ReadDimension(SizeNode, attribute) is uint value ? OfficeLength.Points(value / 20D) : (OfficeLength?)null;

        private void SetLength(string attribute, OfficeLength? value) {
            if (value.HasValue && value.Value.ToEmus() < 0)
                throw new ArgumentOutOfRangeException(nameof(value), "The length cannot be negative.");
            uint? twips = checked((uint?)(value?.ToTwips()));
            OfficeXmlNode size = EnsurePageSize();
            size.SetAttribute(W, attribute, OfficeXmlViewValues.Format(twips));
            size.SetAttribute(W, "code", null);
        }

        private static uint? ReadDimension(OfficeXmlNode size, string attribute) =>
            OfficeXmlViewValues.ParseUInt32(size.Attribute(W, attribute));

        private static ushort? ReadCode(OfficeXmlNode size) =>
            size.Attribute(W, "code") is { } token ? ushort.Parse(token, CultureInfo.InvariantCulture) : (ushort?)null;

        /// <summary>Gets or sets the paper code.</summary>
        public ushort? Code {
            get => ReadCode(SizeNode);
            set => EnsurePageSize().SetAttribute(W, "code", value?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>Gets or sets the section page orientation.</summary>
        public OfficePageOrientation Orientation {
            get => ReadOrientation(SizeNode);
            set {
                _ = value.ToOpenXmlToken();
                OfficeXmlNode section = _section.EditSectionNode();
                OfficeXmlNode size = section.Child(W, "pgSz");
                if (size.IsNone) {
                    size = WordSection.AddSectionChild(section, "pgSz");
                    WritePreset(size, A4);
                }
                ApplyOrientation(size, value);
            }
        }

        private static OfficePageOrientation ReadOrientation(OfficeXmlNode size) {
            string? token = size.Attribute(W, "orient");
            if (token == null) return OfficePageOrientation.Portrait;
            if (WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out OfficePageOrientation orientation)) return orientation;
            throw new FormatException("Unsupported Word page orientation: " + token);
        }

        private static void ApplyOrientation(OfficeXmlNode size, OfficePageOrientation orientation) {
            if (ReadOrientation(size) != orientation) {
                string? width = size.Attribute(W, "w");
                size.SetAttribute(W, "w", size.Attribute(W, "h"));
                size.SetAttribute(W, "h", width);
            }
            size.SetAttribute(W, "orient", orientation.ToOpenXmlToken());
        }

        private static void WritePreset(OfficeXmlNode size, WordPageSizeDefinition preset) {
            size.SetAttribute(W, "w", preset.Width.ToTwips().ToString(CultureInfo.InvariantCulture));
            size.SetAttribute(W, "h", preset.Height.ToTwips().ToString(CultureInfo.InvariantCulture));
            size.SetAttribute(W, "code", preset.PaperCode.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>Initializes the page settings for a section.</summary>
        public WordPageSizes(WordDocument wordDocument, WordSection wordSection) {
            _section = wordSection;
        }
    }
}
