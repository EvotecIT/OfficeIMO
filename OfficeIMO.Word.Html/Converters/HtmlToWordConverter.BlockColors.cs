using SixColor = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static string? ResolveBlockBackground(string styleText, string? inheritedBackground) {
            string? background = null;
            for (int priorityPass = 0; priorityPass < 2; priorityPass++) {
                bool important = priorityPass == 1;
                foreach (string part in styleText.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                    if (!CssStyleMapper.TryParseDeclaration(
                            part,
                            out string name,
                            out string value,
                            out bool declarationIsImportant) ||
                        declarationIsImportant != important ||
                        name != "background-color") {
                        continue;
                    }

                    if (IsCssInheritanceValue(value)) {
                        background = inheritedBackground;
                    } else if (IsCssWideResetValue(value) ||
                               value.Equals("transparent", StringComparison.OrdinalIgnoreCase)) {
                        background = null;
                    } else {
                        CssStyleMapper.CssProperties parsed =
                            CssStyleMapper.ParseStyles("background-color:" + value);
                        if (!string.IsNullOrEmpty(parsed.BackgroundColor)) {
                            double alpha = parsed.BackgroundColorAlpha ?? 1d;
                            background = alpha <= 0d
                                ? null
                                : ResolveOpaqueTextBackground(
                                    parsed.BackgroundColor!,
                                    alpha,
                                    inheritedBackground);
                        }
                    }
                }
            }

            return background;
        }

        private static bool TryResolveBlockBorderColor(
            string value,
            string? backdrop,
            out SixColor color,
            out bool transparent) {
            CssStyleMapper.CssProperties parsed =
                CssStyleMapper.ParseStyles("background-color:" + value);
            if (string.IsNullOrEmpty(parsed.BackgroundColor)) {
                color = SixColor.Black;
                transparent = false;
                return false;
            }

            double alpha = parsed.BackgroundColorAlpha ?? 1d;
            transparent = alpha <= 0d;
            string resolved = transparent
                ? SixColor.Black.ToRgbHex()
                : ResolveOpaqueTextBackground(parsed.BackgroundColor!, alpha, backdrop);
            color = SixColor.Parse("#" + resolved);
            return true;
        }
    }
}
