namespace OfficeIMO.PowerPoint {
    internal static class PowerPointListParagraphDefaults {
        private const double ListHangingPoints = 18D;
        private const double NumberHangingPoints = 28D;
        private const double ListLevelStepPoints = 18D;

        private static readonly PowerPointParagraphStyle DefaultStyle =
            new(lineSpacingMultiplier: 1.15, spaceAfter: global::OfficeIMO.OfficeLength.Points(2));

        internal static void Apply(PowerPointParagraph paragraph) {
            DefaultStyle.Apply(paragraph);
            ApplyDefaultIndent(paragraph);
        }

        private static void ApplyDefaultIndent(PowerPointParagraph paragraph) {
            if ((paragraph.LeftMargin.HasValue || paragraph.Indent.HasValue) &&
                !IsDefaultListIndent(paragraph.LeftMargin?.ToPoints(), paragraph.Indent?.ToPoints())) {
                return;
            }

            int level = paragraph.Level ?? 0;
            if (level < 0) {
                level = 0;
            } else if (level > 8) {
                level = 8;
            }

            // Match the automatic list layout's wider number gutter, including Roman markers.
            double hanging = paragraph.NumberingScheme.HasValue ? NumberHangingPoints : ListHangingPoints;
            paragraph.LeftMargin = global::OfficeIMO.OfficeLength.Points(hanging + (level * ListLevelStepPoints));
            paragraph.Indent = global::OfficeIMO.OfficeLength.Points(-hanging);
        }

        private static bool IsDefaultListIndent(double? leftMarginPoints, double? indentPoints) {
            if (!leftMarginPoints.HasValue || !indentPoints.HasValue) {
                return false;
            }

            double hanging = -indentPoints.Value;
            if (System.Math.Abs(hanging - ListHangingPoints) > 0.000001D &&
                System.Math.Abs(hanging - NumberHangingPoints) > 0.000001D) {
                return false;
            }

            for (int level = 0; level <= 8; level++) {
                double expectedLeft = hanging + (level * ListLevelStepPoints);
                if (System.Math.Abs(leftMarginPoints.Value - expectedLeft) <= 0.000001D) {
                    return true;
                }
            }

            return false;
        }
    }
}
