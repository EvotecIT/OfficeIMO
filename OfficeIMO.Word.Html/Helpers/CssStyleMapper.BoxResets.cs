namespace OfficeIMO.Word.Html {
    internal static partial class CssStyleMapper {
        private static bool IsCssWideBoxReset(string value) {
            string normalized = value.Trim().ToLowerInvariant();
            return normalized is "inherit" or "initial" or "unset" or "revert" or "revert-layer";
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
    }
}
