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
            A.TextAlignmentTypeValues? alignment = null,
            int? level = null,
            double? indentPoints = null,
            double? leftMarginPoints = null,
            double? hangingPoints = null,
            double? lineSpacingPoints = null,
            double? lineSpacingMultiplier = null,
            double? spaceBeforePoints = null,
            double? spaceAfterPoints = null) {
            Alignment = alignment;
            Level = level;
            IndentPoints = indentPoints;
            LeftMarginPoints = leftMarginPoints;
            HangingPoints = hangingPoints;
            LineSpacingPoints = lineSpacingPoints;
            LineSpacingMultiplier = lineSpacingMultiplier;
            SpaceBeforePoints = spaceBeforePoints;
            SpaceAfterPoints = spaceAfterPoints;
        }

        /// <summary>
        /// Paragraph alignment.
        /// </summary>
        public A.TextAlignmentTypeValues? Alignment { get; }

        /// <summary>
        /// Bullet/list level (0-8).
        /// </summary>
        public int? Level { get; }

        /// <summary>
        /// Indentation in points.
        /// </summary>
        public double? IndentPoints { get; }

        /// <summary>
        /// Left margin in points.
        /// </summary>
        public double? LeftMarginPoints { get; }

        /// <summary>
        /// Hanging indent in points.
        /// </summary>
        public double? HangingPoints { get; }

        /// <summary>
        /// Line spacing in points.
        /// </summary>
        public double? LineSpacingPoints { get; }

        /// <summary>
        /// Line spacing multiplier (1.0 = 100%).
        /// </summary>
        public double? LineSpacingMultiplier { get; }

        /// <summary>
        /// Space before the paragraph in points.
        /// </summary>
        public double? SpaceBeforePoints { get; }

        /// <summary>
        /// Space after the paragraph in points.
        /// </summary>
        public double? SpaceAfterPoints { get; }

        /// <summary>
        /// A compact paragraph preset.
        /// </summary>
        public static PowerPointParagraphStyle Compact => new(lineSpacingMultiplier: 1.0, spaceAfterPoints: 0);

        /// <summary>
        /// A comfortable paragraph preset.
        /// </summary>
        public static PowerPointParagraphStyle Comfortable => new(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

        /// <summary>
        /// A spacious paragraph preset.
        /// </summary>
        public static PowerPointParagraphStyle Spacious => new(lineSpacingMultiplier: 1.3, spaceAfterPoints: 4);

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
            if (HangingPoints != null) {
                paragraph.SetHangingPoints(HangingPoints.Value);
            } else {
                if (LeftMarginPoints != null) {
                    paragraph.LeftMarginPoints = LeftMarginPoints;
                }
                if (IndentPoints != null) {
                    paragraph.IndentPoints = IndentPoints;
                }
            }

            if (LineSpacingPoints != null) {
                paragraph.LineSpacingPoints = LineSpacingPoints;
            } else if (LineSpacingMultiplier != null) {
                paragraph.LineSpacingMultiplier = LineSpacingMultiplier;
            }

            if (SpaceBeforePoints != null) {
                paragraph.SpaceBeforePoints = SpaceBeforePoints;
            }
            if (SpaceAfterPoints != null) {
                paragraph.SpaceAfterPoints = SpaceAfterPoints;
            }
        }

        /// <summary>
        /// Returns a copy with a new alignment.
        /// </summary>
        public PowerPointParagraphStyle WithAlignment(A.TextAlignmentTypeValues alignment) {
            return new PowerPointParagraphStyle(alignment, Level, IndentPoints, LeftMarginPoints, HangingPoints,
                LineSpacingPoints, LineSpacingMultiplier, SpaceBeforePoints, SpaceAfterPoints);
        }

        /// <summary>
        /// Returns a copy with a new list level.
        /// </summary>
        public PowerPointParagraphStyle WithLevel(int level) {
            return new PowerPointParagraphStyle(Alignment, level, IndentPoints, LeftMarginPoints, HangingPoints,
                LineSpacingPoints, LineSpacingMultiplier, SpaceBeforePoints, SpaceAfterPoints);
        }

        /// <summary>
        /// Returns a copy with a new indentation value (points).
        /// </summary>
        public PowerPointParagraphStyle WithIndentPoints(double? indentPoints) {
            return new PowerPointParagraphStyle(Alignment, Level, indentPoints, LeftMarginPoints, HangingPoints,
                LineSpacingPoints, LineSpacingMultiplier, SpaceBeforePoints, SpaceAfterPoints);
        }

        /// <summary>
        /// Returns a copy with a new left margin value (points).
        /// </summary>
        public PowerPointParagraphStyle WithLeftMarginPoints(double? leftMarginPoints) {
            return new PowerPointParagraphStyle(Alignment, Level, IndentPoints, leftMarginPoints, HangingPoints,
                LineSpacingPoints, LineSpacingMultiplier, SpaceBeforePoints, SpaceAfterPoints);
        }

        /// <summary>
        /// Returns a copy with a new hanging indent value (points).
        /// </summary>
        public PowerPointParagraphStyle WithHangingPoints(double? hangingPoints) {
            return new PowerPointParagraphStyle(Alignment, Level, IndentPoints, LeftMarginPoints, hangingPoints,
                LineSpacingPoints, LineSpacingMultiplier, SpaceBeforePoints, SpaceAfterPoints);
        }

        /// <summary>
        /// Returns a copy with a new line spacing value (points).
        /// </summary>
        public PowerPointParagraphStyle WithLineSpacingPoints(double? lineSpacingPoints) {
            return new PowerPointParagraphStyle(Alignment, Level, IndentPoints, LeftMarginPoints, HangingPoints,
                lineSpacingPoints, LineSpacingMultiplier, SpaceBeforePoints, SpaceAfterPoints);
        }

        /// <summary>
        /// Returns a copy with a new line spacing multiplier.
        /// </summary>
        public PowerPointParagraphStyle WithLineSpacingMultiplier(double? lineSpacingMultiplier) {
            return new PowerPointParagraphStyle(Alignment, Level, IndentPoints, LeftMarginPoints, HangingPoints,
                LineSpacingPoints, lineSpacingMultiplier, SpaceBeforePoints, SpaceAfterPoints);
        }

        /// <summary>
        /// Returns a copy with a new space before value (points).
        /// </summary>
        public PowerPointParagraphStyle WithSpaceBeforePoints(double? spaceBeforePoints) {
            return new PowerPointParagraphStyle(Alignment, Level, IndentPoints, LeftMarginPoints, HangingPoints,
                LineSpacingPoints, LineSpacingMultiplier, spaceBeforePoints, SpaceAfterPoints);
        }

        /// <summary>
        /// Returns a copy with a new space after value (points).
        /// </summary>
        public PowerPointParagraphStyle WithSpaceAfterPoints(double? spaceAfterPoints) {
            return new PowerPointParagraphStyle(Alignment, Level, IndentPoints, LeftMarginPoints, HangingPoints,
                LineSpacingPoints, LineSpacingMultiplier, SpaceBeforePoints, spaceAfterPoints);
        }
    }
}
