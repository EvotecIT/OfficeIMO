using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordPageSizes {
        /// <summary>Gets a preset's physical dimensions and printer paper code. Unknown returns null.</summary>
        public static WordPageSizeDefinition? GetDefinition(WordPageSize pageSize) => pageSize switch {
            WordPageSize.Unknown => null,
            WordPageSize.Letter => Letter,
            WordPageSize.Legal => Legal,
            WordPageSize.Statement => Statement,
            WordPageSize.Executive => Executive,
            WordPageSize.A3 => A3,
            WordPageSize.A4 => A4,
            WordPageSize.A5 => A5,
            WordPageSize.A6 => A6,
            WordPageSize.B5 => B5,
            WordPageSize.Tabloid => Tabloid,
            WordPageSize.B4Jis => B4Jis,
            WordPageSize.Envelope9 => Envelope9,
            WordPageSize.Envelope10 => Envelope10,
            WordPageSize.CSheet => CSheet,
            WordPageSize.EnvelopeDl => EnvelopeDl,
            WordPageSize.EnvelopeC5 => EnvelopeC5,
            WordPageSize.EnvelopeC4 => EnvelopeC4,
            WordPageSize.EnvelopeB5 => EnvelopeB5,
            WordPageSize.EnvelopeMonarch => EnvelopeMonarch,
            _ => throw new ArgumentOutOfRangeException(nameof(pageSize))
        };

        private static PageSize? GetDefault(WordPageSize? pageSize) =>
            pageSize.HasValue && GetDefinition(pageSize.Value) is { } definition ? ToOpenXmlPageSize(definition) : null;

        /// <summary>A3 paper, 297 × 420 millimeters.</summary>
        public static WordPageSizeDefinition A3 { get; } = FromPhysicalSize(OfficePageSizes.A3, 8);
        /// <summary>A4 paper, 210 × 297 millimeters.</summary>
        public static WordPageSizeDefinition A4 { get; } = FromPhysicalSize(OfficePageSizes.A4, 9);
        /// <summary>A5 paper, 148 × 210 millimeters.</summary>
        public static WordPageSizeDefinition A5 { get; } = FromPhysicalSize(OfficePageSizes.A5, 11);
        /// <summary>A6 paper, 105 × 148 millimeters.</summary>
        public static WordPageSizeDefinition A6 { get; } = FromPhysicalSize(OfficePageSizes.A6, 70);
        /// <summary>Executive paper, 7.25 × 10.5 inches.</summary>
        public static WordPageSizeDefinition Executive { get; } = FromPhysicalSize(OfficePageSizes.Executive, 7);
        /// <summary>JIS B5 paper, 182 × 257 millimeters. This retains the established B5 preset.</summary>
        public static WordPageSizeDefinition B5 { get; } = FromPhysicalSize(OfficePageSizes.B5Jis, 13);
        /// <summary>Statement paper, 5.5 × 8.5 inches.</summary>
        public static WordPageSizeDefinition Statement { get; } = FromPhysicalSize(OfficePageSizes.Statement, 6);
        /// <summary>Legal paper, 8.5 × 14 inches.</summary>
        public static WordPageSizeDefinition Legal { get; } = FromPhysicalSize(OfficePageSizes.Legal, 5);
        /// <summary>Letter paper, 8.5 × 11 inches.</summary>
        public static WordPageSizeDefinition Letter { get; } = FromPhysicalSize(OfficePageSizes.Letter, 1);
        /// <summary>Tabloid paper, 11 × 17 inches.</summary>
        public static WordPageSizeDefinition Tabloid { get; } = FromPhysicalSize(OfficePageSizes.Tabloid, 3);
        /// <summary>JIS B4 paper, 257 × 364 millimeters.</summary>
        public static WordPageSizeDefinition B4Jis { get; } = FromPhysicalSize(OfficePageSizes.B4Jis, 12);
        /// <summary>Number 9 envelope, 3.875 × 8.875 inches.</summary>
        public static WordPageSizeDefinition Envelope9 { get; } = FromPhysicalSize(OfficePageSizes.Envelope9, 19);
        /// <summary>Number 10 envelope, 4.125 × 9.5 inches.</summary>
        public static WordPageSizeDefinition Envelope10 { get; } = FromPhysicalSize(OfficePageSizes.Envelope10, 20);
        /// <summary>C sheet, 17 × 22 inches.</summary>
        public static WordPageSizeDefinition CSheet { get; } = FromPhysicalSize(OfficePageSizes.CSheet, 24);
        /// <summary>DL envelope, 110 × 220 millimeters.</summary>
        public static WordPageSizeDefinition EnvelopeDl { get; } = FromPhysicalSize(OfficePageSizes.EnvelopeDl, 27);
        /// <summary>C5 envelope, 162 × 229 millimeters.</summary>
        public static WordPageSizeDefinition EnvelopeC5 { get; } = FromPhysicalSize(OfficePageSizes.EnvelopeC5, 28);
        /// <summary>C4 envelope, 229 × 324 millimeters.</summary>
        public static WordPageSizeDefinition EnvelopeC4 { get; } = FromPhysicalSize(OfficePageSizes.EnvelopeC4, 30);
        /// <summary>B5 envelope, 176 × 250 millimeters.</summary>
        public static WordPageSizeDefinition EnvelopeB5 { get; } = FromPhysicalSize(OfficePageSizes.EnvelopeB5, 34);
        /// <summary>Monarch envelope, 3.875 × 7.5 inches.</summary>
        public static WordPageSizeDefinition EnvelopeMonarch { get; } = FromPhysicalSize(OfficePageSizes.EnvelopeMonarch, 37);

        private static WordPageSizeDefinition FromPhysicalSize(OfficePageSize size, ushort paperCode) =>
            new WordPageSizeDefinition(
                checked((uint)Math.Round(size.ToPointWidth() * 20D, MidpointRounding.AwayFromZero)),
                checked((uint)Math.Round(size.ToPointHeight() * 20D, MidpointRounding.AwayFromZero)), paperCode);
    }
}
