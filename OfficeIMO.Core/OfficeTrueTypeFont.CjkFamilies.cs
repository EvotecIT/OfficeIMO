using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing {
    public sealed partial class OfficeTrueTypeFont {
        // Unembedded Adobe CJK faces need a substitute that covers their script rather
        // than the Latin bitmap fallback. These are installed-font candidates only;
        // no font files are bundled or required. PDF retains its substitution diagnostic.
        private static string[]? CjkSubstituteFamilies(string key) => key switch {
            "heiseimin" => new[] { "Yu Mincho", "MS Mincho", "Songti SC", "Arial Unicode MS", "Noto Serif CJK JP", "Noto Serif JP" },
            "heiseikakugo" => new[] { "Yu Gothic", "MS Gothic", "Hiragino Sans GB", "Arial Unicode MS", "Noto Sans CJK JP", "Noto Sans JP" },
            "stsong" => new[] { "Songti SC", "SimSun", "Arial Unicode MS", "Noto Serif CJK SC", "Noto Serif SC" },
            "msung" => new[] { "Songti TC", "PMingLiU", "MingLiU", "Arial Unicode MS", "Noto Serif CJK TC", "Noto Serif TC" },
            "hysmyeongjo" => new[] { "Batang", "Arial Unicode MS", "Noto Serif CJK KR", "Noto Serif KR" },
            "hygothic" => new[] { "Gulim", "Arial Unicode MS", "Noto Sans CJK KR", "Noto Sans KR" },
            _ => null
        };

        private static IEnumerable<string> CandidateCjkFamilyPaths(string key) {
            if (key == "songtisc" || key == "songtitc") {
                yield return "/System/Library/Fonts/Supplemental/Songti.ttc";
            } else if (key == "arialunicodems") {
                yield return "/System/Library/Fonts/Supplemental/Arial Unicode.ttf";
                yield return "/Library/Fonts/Arial Unicode.ttf";
            }

            string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrEmpty(windows)) {
                string fonts = Path.Combine(windows, "Fonts");
                string? file = key switch {
                    "yumincho" => "yumin.ttf",
                    "msmincho" => "msmincho.ttc",
                    "yugothic" => "YuGothR.ttc",
                    "msgothic" => "msgothic.ttc",
                    "simsun" => "simsun.ttc",
                    "pmingliu" or "mingliu" => "mingliu.ttc",
                    "batang" => "batang.ttc",
                    "gulim" => "gulim.ttc",
                    "arialunicodems" => "ARIALUNI.TTF",
                    _ => null
                };
                if (file != null) yield return Path.Combine(fonts, file);
            }

            // Common optional Noto installations remain subject to the ordinary font
            // loader's outline-format and family-name checks.
            string? noto = key switch {
                "notoserifcjkjp" or "notoserifcjksc" or "notoserifcjktc" or "notoserifcjkkr" => "NotoSerifCJK-Regular.ttc",
                "notosanscjkjp" or "notosanscjksc" or "notosanscjktc" or "notosanscjkkr" => "NotoSansCJK-Regular.ttc",
                _ => null
            };
            if (noto != null) {
                yield return Path.Combine("/usr/share/fonts/opentype/noto", noto);
                yield return Path.Combine("/usr/share/fonts/truetype/noto", noto);
            }
        }
    }
}
