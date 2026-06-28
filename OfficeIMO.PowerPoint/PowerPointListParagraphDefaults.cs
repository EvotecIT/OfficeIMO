namespace OfficeIMO.PowerPoint {
    internal static class PowerPointListParagraphDefaults {
        private const double ListHangingPoints = 18D;
        private const double ListLevelStepPoints = 18D;

        private static readonly PowerPointParagraphStyle DefaultStyle =
            new(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

        internal static void Apply(PowerPointParagraph paragraph) {
            DefaultStyle.Apply(paragraph);
            ApplyDefaultIndent(paragraph);
        }

        private static void ApplyDefaultIndent(PowerPointParagraph paragraph) {
            if ((paragraph.LeftMarginPoints.HasValue || paragraph.IndentPoints.HasValue) &&
                !IsDefaultListIndent(paragraph.LeftMarginPoints, paragraph.IndentPoints)) {
                return;
            }

            int level = paragraph.Level ?? 0;
            if (level < 0) {
                level = 0;
            } else if (level > 8) {
                level = 8;
            }

            paragraph.LeftMarginPoints = ListHangingPoints + (level * ListLevelStepPoints);
            paragraph.IndentPoints = -ListHangingPoints;
        }

        private static bool IsDefaultListIndent(double? leftMarginPoints, double? indentPoints) {
            if (!leftMarginPoints.HasValue || !indentPoints.HasValue) {
                return false;
            }

            if (System.Math.Abs(indentPoints.Value + ListHangingPoints) > 0.000001D) {
                return false;
            }

            for (int level = 0; level <= 8; level++) {
                double expectedLeft = ListHangingPoints + (level * ListLevelStepPoints);
                if (System.Math.Abs(leftMarginPoints.Value - expectedLeft) <= 0.000001D) {
                    return true;
                }
            }

            return false;
        }
    }
}
