using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed class RichSeg {
        public RichSeg(
            string text,
            bool bold,
            bool italic,
            bool underline,
            bool strike,
            PdfColor? color,
            PdfColor? backgroundColor,
            string? uri,
            string? destinationName,
            string? contents,
            PdfStandardFont font,
            double fontSize,
            PdfTextBaseline baseline,
            double measuredWidth,
            bool leadingSpace = false,
            double leadingAdvance = 0,
            bool leadingSpaceIsExpandable = true,
            PdfTabLeaderStyle leadingTabLeader = PdfTabLeaderStyle.None,
            bool endsWithHardBreak = false,
            bool endsWithTextSeparator = false,
            PdfInlineElement? inlineElement = null,
            PdfNamedFontFace? namedFont = null,
            OfficeIMO.Drawing.OfficeTextDecorationStyle underlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None,
            OfficeIMO.Drawing.OfficeTextDecorationStyle strikeStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None,
            PdfColor? decorationColor = null,
            OfficeTextFeatureSettings? featureSettings = null,
            OfficeTextDirection textDirection = OfficeTextDirection.Auto,
            double fontMetricScale = 1D,
            PdfTabStop? leadingTabStop = null,
            OfficeIMO.Drawing.OfficeTextDecorationStyle leadingUnderlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None,
            PdfColor? leadingDecorationColor = null,
            double leadingDecorationFontSize = 0,
            double leadingDecorationTextRise = 0) {
            Text = text;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strike = strike;
            Color = color;
            BackgroundColor = backgroundColor;
            Uri = uri;
            DestinationName = destinationName;
            Contents = contents;
            Font = font;
            FontSize = fontSize;
            Baseline = baseline;
            MeasuredWidth = measuredWidth;
            LeadingSpace = leadingSpace;
            LeadingAdvance = leadingAdvance;
            LeadingTabStop = leadingTabStop;
            LeadingUnderlineStyle = leadingUnderlineStyle;
            LeadingDecorationColor = leadingDecorationColor;
            LeadingDecorationFontSize = leadingDecorationFontSize;
            LeadingDecorationTextRise = leadingDecorationTextRise;
            LeadingSpaceIsExpandable = leadingSpaceIsExpandable;
            LeadingTabLeader = leadingTabLeader;
            EndsWithHardBreak = endsWithHardBreak;
            EndsWithTextSeparator = endsWithTextSeparator;
            InlineElement = inlineElement;
            NamedFont = namedFont;
            UnderlineStyle = underlineStyle != OfficeIMO.Drawing.OfficeTextDecorationStyle.None
                ? underlineStyle
                : underline ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            StrikeStyle = strikeStyle != OfficeIMO.Drawing.OfficeTextDecorationStyle.None
                ? strikeStyle
                : strike ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            DecorationColor = decorationColor;
            FeatureSettings = featureSettings ?? OfficeTextFeatureSettings.Default;
            TextDirection = textDirection;
            FontMetricScale = fontMetricScale;
        }

        public string Text { get; }

        public bool Bold { get; }

        public bool Italic { get; }

        public bool Underline { get; }

        public OfficeIMO.Drawing.OfficeTextDecorationStyle UnderlineStyle { get; }

        public bool Strike { get; }

        public OfficeIMO.Drawing.OfficeTextDecorationStyle StrikeStyle { get; }

        public PdfColor? Color { get; }

        public PdfColor? BackgroundColor { get; }

        public PdfColor? DecorationColor { get; }

        public string? Uri { get; }

        public string? DestinationName { get; }

        public string? Contents { get; }

        public PdfStandardFont Font { get; }

        public double FontSize { get; }

        public PdfTextBaseline Baseline { get; }

        public double MeasuredWidth { get; }

        public bool LeadingSpace { get; }

        public double LeadingAdvance { get; }
        public PdfTabStop? LeadingTabStop { get; }
        public OfficeIMO.Drawing.OfficeTextDecorationStyle LeadingUnderlineStyle { get; }
        public PdfColor? LeadingDecorationColor { get; }
        public double LeadingDecorationFontSize { get; }
        public double LeadingDecorationTextRise { get; }

        public bool LeadingSpaceIsExpandable { get; }

        public PdfTabLeaderStyle LeadingTabLeader { get; }

        public bool EndsWithHardBreak { get; }

        public bool EndsWithTextSeparator { get; }

        public PdfInlineElement? InlineElement { get; }

        public PdfNamedFontFace? NamedFont { get; }

        public OfficeTextFeatureSettings FeatureSettings { get; }

        public OfficeTextDirection TextDirection { get; }
        public double FontMetricScale { get; }

        public RichSeg WithEndsWithHardBreak() =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, Uri, DestinationName, Contents, Font, FontSize, Baseline, MeasuredWidth, LeadingSpace, LeadingAdvance, LeadingSpaceIsExpandable, LeadingTabLeader, true, true, InlineElement, NamedFont, UnderlineStyle, StrikeStyle, DecorationColor, FeatureSettings, TextDirection, FontMetricScale, LeadingTabStop, LeadingUnderlineStyle, LeadingDecorationColor, LeadingDecorationFontSize, LeadingDecorationTextRise);

        public RichSeg WithEndsWithTextSeparator() =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, Uri, DestinationName, Contents, Font, FontSize, Baseline, MeasuredWidth, LeadingSpace, LeadingAdvance, LeadingSpaceIsExpandable, LeadingTabLeader, EndsWithHardBreak, true, InlineElement, NamedFont, UnderlineStyle, StrikeStyle, DecorationColor, FeatureSettings, TextDirection, FontMetricScale, LeadingTabStop, LeadingUnderlineStyle, LeadingDecorationColor, LeadingDecorationFontSize, LeadingDecorationTextRise);

        public RichSeg WithoutLink() =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, null, null, null, Font, FontSize, Baseline, MeasuredWidth, LeadingSpace, LeadingAdvance, LeadingSpaceIsExpandable, LeadingTabLeader, EndsWithHardBreak, EndsWithTextSeparator, InlineElement, NamedFont, UnderlineStyle, StrikeStyle, DecorationColor, FeatureSettings, TextDirection, FontMetricScale, LeadingTabStop, LeadingUnderlineStyle, LeadingDecorationColor, LeadingDecorationFontSize, LeadingDecorationTextRise);

        public RichSeg WithLeadingAdvance(double advance) =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, Uri, DestinationName, Contents, Font, FontSize, Baseline, MeasuredWidth, advance > 0, advance, LeadingSpaceIsExpandable, LeadingTabLeader, EndsWithHardBreak, EndsWithTextSeparator, InlineElement, NamedFont, UnderlineStyle, StrikeStyle, DecorationColor, FeatureSettings, TextDirection, FontMetricScale, LeadingTabStop, LeadingUnderlineStyle, LeadingDecorationColor, LeadingDecorationFontSize, LeadingDecorationTextRise);
    }

}
