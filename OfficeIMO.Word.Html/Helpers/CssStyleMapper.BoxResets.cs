namespace OfficeIMO.Word.Html {
    internal static partial class CssStyleMapper {
        private static bool IsCssWideBoxReset(string value) {
            string normalized = value.Trim().ToLowerInvariant();
            return normalized is "initial" or "unset" or "revert" or "revert-layer";
        }

        private static bool IsCssBoxInheritance(string value) =>
            value.Trim().Equals("inherit", StringComparison.OrdinalIgnoreCase);

        private static void CopyMargin(CssProperties? source, CssProperties result) {
            result.MarginTop = source?.MarginTop;
            result.MarginRight = source?.MarginRight;
            result.MarginBottom = source?.MarginBottom;
            result.MarginLeft = source?.MarginLeft;
        }

        private static void CopyPadding(CssProperties? source, CssProperties result) {
            result.PaddingTop = source?.PaddingTop;
            result.PaddingRight = source?.PaddingRight;
            result.PaddingBottom = source?.PaddingBottom;
            result.PaddingLeft = source?.PaddingLeft;
        }

        private static void ResetMargin(CssProperties result) {
            result.MarginTop = null;
            result.MarginRight = null;
            result.MarginBottom = null;
            result.MarginLeft = null;
        }

        private static void ResetPadding(CssProperties result) {
            result.PaddingTop = null;
            result.PaddingRight = null;
            result.PaddingBottom = null;
            result.PaddingLeft = null;
        }

        private static void ResetLogicalBoxProperty(
            string name,
            string prefix,
            bool rightToLeft,
            CssProperties result) {
            bool margin = prefix.Equals("margin", StringComparison.OrdinalIgnoreCase);
            void ResetInlineStart() {
                if (margin) {
                    if (rightToLeft) result.MarginRight = null;
                    else result.MarginLeft = null;
                } else {
                    if (rightToLeft) result.PaddingRight = null;
                    else result.PaddingLeft = null;
                }
            }
            void ResetInlineEnd() {
                if (margin) {
                    if (rightToLeft) result.MarginLeft = null;
                    else result.MarginRight = null;
                } else {
                    if (rightToLeft) result.PaddingLeft = null;
                    else result.PaddingRight = null;
                }
            }
            void ResetBlockStart() {
                if (margin) result.MarginTop = null;
                else result.PaddingTop = null;
            }
            void ResetBlockEnd() {
                if (margin) result.MarginBottom = null;
                else result.PaddingBottom = null;
            }

            if (name.EndsWith("-inline", StringComparison.Ordinal)) {
                ResetInlineStart();
                ResetInlineEnd();
            } else if (name.EndsWith("-inline-start", StringComparison.Ordinal)) {
                ResetInlineStart();
            } else if (name.EndsWith("-inline-end", StringComparison.Ordinal)) {
                ResetInlineEnd();
            } else if (name.EndsWith("-block", StringComparison.Ordinal)) {
                ResetBlockStart();
                ResetBlockEnd();
            } else if (name.EndsWith("-block-start", StringComparison.Ordinal)) {
                ResetBlockStart();
            } else if (name.EndsWith("-block-end", StringComparison.Ordinal)) {
                ResetBlockEnd();
            }
        }

        private static void InheritLogicalBoxProperty(
            string name,
            string prefix,
            bool rightToLeft,
            CssProperties result,
            CssProperties? inheritedBox,
            bool inheritedRightToLeft) {
            bool margin = prefix.Equals("margin", StringComparison.OrdinalIgnoreCase);
            int? inheritedInlineStart = margin
                ? inheritedRightToLeft ? inheritedBox?.MarginRight : inheritedBox?.MarginLeft
                : inheritedRightToLeft ? inheritedBox?.PaddingRight : inheritedBox?.PaddingLeft;
            int? inheritedInlineEnd = margin
                ? inheritedRightToLeft ? inheritedBox?.MarginLeft : inheritedBox?.MarginRight
                : inheritedRightToLeft ? inheritedBox?.PaddingLeft : inheritedBox?.PaddingRight;
            int? inheritedBlockStart = margin ? inheritedBox?.MarginTop : inheritedBox?.PaddingTop;
            int? inheritedBlockEnd = margin ? inheritedBox?.MarginBottom : inheritedBox?.PaddingBottom;

            void SetInlineStart(int? value) {
                if (margin) {
                    if (rightToLeft) result.MarginRight = value;
                    else result.MarginLeft = value;
                } else {
                    if (rightToLeft) result.PaddingRight = value;
                    else result.PaddingLeft = value;
                }
            }
            void SetInlineEnd(int? value) {
                if (margin) {
                    if (rightToLeft) result.MarginLeft = value;
                    else result.MarginRight = value;
                } else {
                    if (rightToLeft) result.PaddingLeft = value;
                    else result.PaddingRight = value;
                }
            }
            void SetBlockStart(int? value) {
                if (margin) result.MarginTop = value;
                else result.PaddingTop = value;
            }
            void SetBlockEnd(int? value) {
                if (margin) result.MarginBottom = value;
                else result.PaddingBottom = value;
            }

            if (name.EndsWith("-inline", StringComparison.Ordinal)) {
                SetInlineStart(inheritedInlineStart);
                SetInlineEnd(inheritedInlineEnd);
            } else if (name.EndsWith("-inline-start", StringComparison.Ordinal)) {
                SetInlineStart(inheritedInlineStart);
            } else if (name.EndsWith("-inline-end", StringComparison.Ordinal)) {
                SetInlineEnd(inheritedInlineEnd);
            } else if (name.EndsWith("-block", StringComparison.Ordinal)) {
                SetBlockStart(inheritedBlockStart);
                SetBlockEnd(inheritedBlockEnd);
            } else if (name.EndsWith("-block-start", StringComparison.Ordinal)) {
                SetBlockStart(inheritedBlockStart);
            } else if (name.EndsWith("-block-end", StringComparison.Ordinal)) {
                SetBlockEnd(inheritedBlockEnd);
            }
        }
    }
}
