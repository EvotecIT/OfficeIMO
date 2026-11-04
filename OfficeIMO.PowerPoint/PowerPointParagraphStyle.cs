using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Defines paragraph formatting that can be applied to text boxes.
    /// </summary>
    public readonly struct PowerPointParagraphStyle {
        /// <summary>
        /// Creates a new paragraph style instance.
        /// </summary>
        public PowerPointParagraphStyle(
            PowerPointTextAlignment? alignment = null,
            int? level = null,
            global::OfficeIMO.OfficeLength? indent = null,
            global::OfficeIMO.OfficeLength? leftMargin = null,
            global::OfficeIMO.OfficeLength? hanging = null,
            global::OfficeIMO.OfficeLength? lineSpacing = null,
            double? lineSpacingMultiplier = null,
            global::OfficeIMO.OfficeLength? spaceBefore = null,
            global::OfficeIMO.OfficeLength? spaceAfter = null) {
            Alignment = alignment;
            Level = level;
            Indent = indent;
            LeftMargin = leftMargin;
            Hanging = hanging;
            LineSpacing = lineSpacing;
            LineSpacingMultiplier = lineSpacingMultiplier;
            SpaceBefore = spaceBefore;
            SpaceAfter = spaceAfter;
        }

        /// <summary>
        /// Paragraph alignment.
        /// </summary>
        public PowerPointTextAlignment? Alignment { get; }

        /// <summary>
        /// Bullet/list level (0-8).
        /// </summary>
        public int? Level { get; }

        /// <summary>
        /// Indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? Indent { get; }

        /// <summary>
        /// Left margin as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LeftMargin { get; }

        /// <summary>
        /// Hanging indent as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? Hanging { get; }

        /// <summary>
        /// Line spacing as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacing { get; }

        /// <summary>
        /// Line spacing multiplier (1.0 = 100%).
        /// </summary>
        public double? LineSpacingMultiplier { get; }

        /// <summary>
        /// Space before the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? SpaceBefore { get; }

        /// <summary>
        /// Space after the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? SpaceAfter { get; }

        /// <summary>
        /// A compact paragraph preset.
        /// </summary>
        public static PowerPointParagraphStyle Compact => new(lineSpacingMultiplier: 1.0, spaceAfter: global::OfficeIMO.OfficeLength.Points(0));

        /// <summary>
        /// A comfortable paragraph preset.
        /// </summary>
        public static PowerPointParagraphStyle Comfortable => new(lineSpacingMultiplier: 1.15, spaceAfter: global::OfficeIMO.OfficeLength.Points(2));

        /// <summary>
        /// A spacious paragraph preset.
        /// </summary>
        public static PowerPointParagraphStyle Spacious => new(lineSpacingMultiplier: 1.3, spaceAfter: global::OfficeIMO.OfficeLength.Points(4));

        /// <summary>
        /// Applies the style to a paragraph.
        /// </summary>
        public void Apply(PowerPointParagraph paragraph) {
            if (paragraph == null) {
                throw new ArgumentNullException(nameof(paragraph));
            }

            if (Alignment != null) {
                paragraph.Alignment = Alignment;
            }
            if (Level != null) {
                paragraph.Level = Level;
            }
            if (Hanging != null) {
                paragraph.SetHanging(Hanging.Value);
            } else {
                if (LeftMargin != null) {
                    paragraph.LeftMargin = LeftMargin;
                }
                if (Indent != null) {
                    paragraph.Indent = Indent;
                }
            }

            if (LineSpacing != null) {
                paragraph.LineSpacing = LineSpacing;
            } else if (LineSpacingMultiplier != null) {
                paragraph.LineSpacingMultiplier = LineSpacingMultiplier;
            }

            if (SpaceBefore != null) {
                paragraph.SpaceBefore = SpaceBefore;
            }
            if (SpaceAfter != null) {
                paragraph.SpaceAfter = SpaceAfter;
            }
        }

        /// <summary>
        /// Returns a copy with a new alignment.
        /// </summary>
        public PowerPointParagraphStyle WithAlignment(PowerPointTextAlignment alignment) {
            return new PowerPointParagraphStyle(alignment, Level, Indent, LeftMargin, Hanging,
                LineSpacing, LineSpacingMultiplier, SpaceBefore, SpaceAfter);
        }

        /// <summary>
        /// Returns a copy with a new list level.
        /// </summary>
        public PowerPointParagraphStyle WithLevel(int level) {
            return new PowerPointParagraphStyle(Alignment, level, Indent, LeftMargin, Hanging,
                LineSpacing, LineSpacingMultiplier, SpaceBefore, SpaceAfter);
        }

        /// <summary>
        /// Returns a copy with a new indentation physical length.
        /// </summary>
        public PowerPointParagraphStyle WithIndent(global::OfficeIMO.OfficeLength? indent) {
            return new PowerPointParagraphStyle(Alignment, Level, indent, LeftMargin, Hanging,
                LineSpacing, LineSpacingMultiplier, SpaceBefore, SpaceAfter);
        }

        /// <summary>
        /// Returns a copy with a new left margin physical length.
        /// </summary>
        public PowerPointParagraphStyle WithLeftMargin(global::OfficeIMO.OfficeLength? leftMargin) {
            return new PowerPointParagraphStyle(Alignment, Level, Indent, leftMargin, Hanging,
                LineSpacing, LineSpacingMultiplier, SpaceBefore, SpaceAfter);
        }

        /// <summary>
        /// Returns a copy with a new hanging indent physical length.
        /// </summary>
        public PowerPointParagraphStyle WithHanging(global::OfficeIMO.OfficeLength? hanging) {
            return new PowerPointParagraphStyle(Alignment, Level, Indent, LeftMargin, hanging,
                LineSpacing, LineSpacingMultiplier, SpaceBefore, SpaceAfter);
        }

        /// <summary>
        /// Returns a copy with a new line spacing physical length.
        /// </summary>
        public PowerPointParagraphStyle WithLineSpacing(global::OfficeIMO.OfficeLength? lineSpacing) {
            return new PowerPointParagraphStyle(Alignment, Level, Indent, LeftMargin, Hanging,
                lineSpacing, LineSpacingMultiplier, SpaceBefore, SpaceAfter);
        }

        /// <summary>
        /// Returns a copy with a new line spacing multiplier.
        /// </summary>
        public PowerPointParagraphStyle WithLineSpacingMultiplier(double? lineSpacingMultiplier) {
            return new PowerPointParagraphStyle(Alignment, Level, Indent, LeftMargin, Hanging,
                LineSpacing, lineSpacingMultiplier, SpaceBefore, SpaceAfter);
        }

        /// <summary>
        /// Returns a copy with a new space before physical length.
        /// </summary>
        public PowerPointParagraphStyle WithSpaceBefore(global::OfficeIMO.OfficeLength? spaceBefore) {
            return new PowerPointParagraphStyle(Alignment, Level, Indent, LeftMargin, Hanging,
                LineSpacing, LineSpacingMultiplier, spaceBefore, SpaceAfter);
        }

        /// <summary>
        /// Returns a copy with a new space after physical length.
        /// </summary>
        public PowerPointParagraphStyle WithSpaceAfter(global::OfficeIMO.OfficeLength? spaceAfter) {
            return new PowerPointParagraphStyle(Alignment, Level, Indent, LeftMargin, Hanging,
                LineSpacing, LineSpacingMultiplier, SpaceBefore, spaceAfter);
        }
    }
}
