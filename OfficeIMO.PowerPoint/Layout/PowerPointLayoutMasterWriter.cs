using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Writes the theme, the default text style, the slide master (placeholders and text styles) and the
    ///     built-in slide layouts of a new presentation from an <see cref="OfficeTheme" />. Every run that has no
    ///     explicit formatting, including table cells and placeholders, then resolves to the theme fonts, sizes and
    ///     colours in every renderer, instead of each renderer's own default.
    /// </summary>
    internal static class PowerPointLayoutMasterWriter {
        internal const string NsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
        internal const string NsP = "http://schemas.openxmlformats.org/presentationml/2006/main";
        internal const string NsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        internal const uint FooterIndex = 11U;
        internal const uint SlideNumberIndex = 12U;
        internal const uint DateIndex = 10U;

        internal static void Apply(PresentationPart presentationPart, PowerPointLayoutTheme theme) {
            if (presentationPart.Presentation == null) return;
            P.SlideSize? size = presentationPart.Presentation.SlideSize;
            double width = (size?.Cx?.Value ?? 12192000) / 12700D;
            double height = (size?.Cy?.Value ?? 6858000) / 12700D;
            var grid = new PowerPointLayoutGrid(width, height, theme.MarginLeft, theme.MarginRight, theme.MarginTop);

            presentationPart.Presentation.DefaultTextStyle = new P.DefaultTextStyle(DefaultTextStyleXml(theme));

            foreach (SlideMasterPart masterPart in presentationPart.SlideMasterParts) {
                if (masterPart.ThemePart != null) WriteTheme(masterPart.ThemePart, theme);
                WriteMaster(masterPart, theme, grid);
                foreach (SlideLayoutPart layoutPart in masterPart.SlideLayoutParts) {
                    WriteLayout(layoutPart, theme, grid);
                }
            }

            presentationPart.Presentation.Save();
        }

        // ---------------------------------------------------------------- theme part

        internal static void WriteTheme(ThemePart themePart, PowerPointLayoutTheme theme) {
            A.Theme? root = themePart.Theme;
            if (root == null) return;
            A.ThemeElements? elements = root.ThemeElements;
            A.ColorScheme? colors = elements?.ColorScheme;
            if (colors != null) {
                SetSchemeColor(colors.Dark1Color, theme.Text);
                SetSchemeColor(colors.Light1Color, theme.Background);
                SetSchemeColor(colors.Dark2Color, theme.TextMuted);
                SetSchemeColor(colors.Light2Color, theme.Surface);
                SetSchemeColor(colors.Accent1Color, theme.Accent1);
                SetSchemeColor(colors.Accent2Color, theme.Accent2);
                SetSchemeColor(colors.Accent3Color, theme.Accent3);
                SetSchemeColor(colors.Accent4Color, theme.Accent4);
                SetSchemeColor(colors.Accent5Color, theme.Accent5);
                SetSchemeColor(colors.Accent6Color, theme.Accent6);
                SetSchemeColor(colors.Hyperlink, theme.Accent1);
                SetSchemeColor(colors.FollowedHyperlinkColor, theme.Accent5);
            }

            A.FontScheme? fonts = elements?.FontScheme;
            if (fonts != null) {
                SetSchemeFont(fonts.MajorFont, theme.HeadingFont);
                SetSchemeFont(fonts.MinorFont, theme.BodyFont);
            }

            root.Save();
        }

        private static void SetSchemeColor(OpenXmlCompositeElement? slot, string hex) {
            if (slot == null) return;
            // Keep a system colour (windowText, window) when it already resolves to the requested value.
            if (slot.GetFirstChild<A.SystemColor>() is { } system &&
                string.Equals(system.LastColor?.Value, hex, StringComparison.OrdinalIgnoreCase)) return;
            slot.RemoveAllChildren();
            slot.AppendChild(new A.RgbColorModelHex { Val = hex });
        }

        private static void SetSchemeFont(OpenXmlCompositeElement? font, string typeface) {
            if (font == null) return;
            A.LatinFont latin = font.GetFirstChild<A.LatinFont>() ?? font.PrependChild(new A.LatinFont());
            PowerPointFontClassification.Apply(latin, typeface);
        }

        // ---------------------------------------------------------------- text styles

        private static string Pt(double points) => ((long)Math.Round(points * 12700D)).ToString(CultureInfo.InvariantCulture);
        // Paragraph spacing in whole 1/8 points (the binary PowerPoint field) expressed in 1/100 points.
        private static string Spc(double points) => ((int)Math.Round(points * 4D) * 25).ToString(CultureInfo.InvariantCulture);
        private static string Sz(double points) => ((int)Math.Round(points * 100D)).ToString(CultureInfo.InvariantCulture);
        // Whole percentages: binary PPT stores spacing in whole percent, so a deck saved as .ppt keeps the same spacing.
        private static string Pct(double fraction) => ((int)Math.Round(fraction * 100D) * 1000).ToString(CultureInfo.InvariantCulture);

        /// <summary>PowerPoint's 100 % line spacing is about 1.2 em; the design line heights are CSS-like em multiples.</summary>
        internal static double LineSpacingFraction(double lineHeight) => Math.Round(lineHeight / 1.2D * 100D) / 100D;

        // Master runs stay inside the binary PowerPoint character contract (size, weight, fill, Latin and East
        // Asian typefaces) so a new deck can still be saved as .ppt without loss. Tracking is applied on slide runs.
        private static string Run(string font, double size, bool bold, string color, double tracking = 0D) {
            _ = tracking;
            var b = new StringBuilder();
            b.Append("<a:defRPr sz=\"").Append(Sz(Math.Round(size))).Append('"');
            b.Append(" b=\"").Append(bold ? '1' : '0').Append("\">");
            b.Append(color);
            b.Append("<a:latin typeface=\"").Append(font).Append("\"/>");
            b.Append("<a:ea typeface=\"").Append(font == "+mj-lt" ? "+mj-ea" : "+mn-ea").Append("\"/>");
            b.Append("</a:defRPr>");
            return b.ToString();
        }

        private static string Scheme(string val) => "<a:solidFill><a:schemeClr val=\"" + val + "\"/></a:solidFill>";

        private static string DefaultTextStyleXml(PowerPointLayoutTheme theme) {
            double body = theme[PowerPointTextRole.Body].Size;
            var b = new StringBuilder();
            b.Append("<p:defaultTextStyle xmlns:p=\"").Append(NsP).Append("\" xmlns:a=\"").Append(NsA).Append("\">");
            b.Append("");
            for (int level = 1; level <= 5; level++) {
                b.Append("<a:lvl").Append(level).Append("pPr marL=\"").Append(Pt((level - 1) * 36D))
                    .Append("\" algn=\"l\" defTabSz=\"").Append(Pt(72)).Append("\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">");
                b.Append(Run("+mn-lt", body, false, Scheme("tx1")));
                b.Append("</a:lvl").Append(level).Append("pPr>");
            }
            b.Append("</p:defaultTextStyle>");
            return b.ToString();
        }

        internal static string TextStylesXml(PowerPointLayoutTheme theme) {
            PowerPointTextStyleSpec title = theme[PowerPointTextRole.Title];
            PowerPointTextStyleSpec body = theme[PowerPointTextRole.Body];
            PowerPointTextStyleSpec small = theme[PowerPointTextRole.BodySmall];
            double after = theme.ParagraphAfter;
            var b = new StringBuilder();
            b.Append("<p:txStyles xmlns:p=\"").Append(NsP).Append("\" xmlns:a=\"").Append(NsA).Append("\">");

            b.Append("<p:titleStyle><a:lvl1pPr algn=\"l\" defTabSz=\"").Append(Pt(72)).Append("\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">");
            b.Append("<a:lnSpc><a:spcPct val=\"").Append(Pct(LineSpacingFraction(title.LineHeight))).Append("\"/></a:lnSpc>");
            b.Append("<a:spcBef><a:spcPct val=\"0\"/></a:spcBef><a:buNone/>");
            b.Append(Run("+mj-lt", title.Size, true, Scheme("tx1"), title.Tracking));
            b.Append("</a:lvl1pPr></p:titleStyle>");

            b.Append("<p:bodyStyle>");
            for (int level = 1; level <= 5; level++) {
                bool first = level == 1;
                double size = first ? body.Size : small.Size;
                double marL = first ? theme.ListIndent : theme.ListIndent + 22D * (level - 1);
                b.Append("<a:lvl").Append(level).Append("pPr marL=\"").Append(Pt(marL)).Append("\" indent=\"").Append(Pt(-(first ? theme.ListIndent : 18D)))
                    .Append("\" algn=\"l\" defTabSz=\"").Append(Pt(72)).Append("\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">");
                b.Append("<a:lnSpc><a:spcPct val=\"").Append(Pct(LineSpacingFraction(body.LineHeight))).Append("\"/></a:lnSpc>");
                b.Append("<a:spcBef><a:spcPct val=\"0\"/></a:spcBef>");
                b.Append("<a:spcAft><a:spcPts val=\"").Append(Spc(after * size / body.Size)).Append("\"/></a:spcAft>");
                if (first) {
                    b.Append("<a:buClr><a:schemeClr val=\"accent1\"/></a:buClr><a:buSzPct val=\"100000\"/>");
                    b.Append("<a:buFont typeface=\"Arial\"/><a:buChar char=\"▪\"/>");
                } else {
                    b.Append("<a:buClr><a:srgbClr val=\"").Append(theme.Hairline).Append("\"/></a:buClr><a:buSzPct val=\"100000\"/>");
                    b.Append("<a:buFont typeface=\"Arial\"/><a:buChar char=\"–\"/>");
                }
                b.Append(Run("+mn-lt", size, false, first ? Scheme("tx1") : Scheme("tx2")));
                b.Append("</a:lvl").Append(level).Append("pPr>");
            }
            b.Append("</p:bodyStyle>");

            b.Append("<p:otherStyle>");
            for (int level = 1; level <= 5; level++) {
                b.Append("<a:lvl").Append(level).Append("pPr marL=\"").Append(Pt((level - 1) * 36D))
                    .Append("\" algn=\"l\" defTabSz=\"").Append(Pt(72)).Append("\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">");
                b.Append(Run("+mn-lt", body.Size, false, Scheme("tx1")));
                b.Append("</a:lvl").Append(level).Append("pPr>");
            }
            b.Append("</p:otherStyle>");
            b.Append("</p:txStyles>");
            return b.ToString();
        }

        // ---------------------------------------------------------------- master and layouts

        private sealed class Placeholder {
            internal Placeholder(string name, string? type, uint? index, PowerPointLayoutRect box, string anchor = "t",
                string? size = null, string? listStyle = null, string? text = null, string? orientation = null) {
                Name = name;
                Type = type;
                Index = index;
                Box = box;
                Anchor = anchor;
                Size = size;
                ListStyle = listStyle;
                Text = text;
                Orientation = orientation;
            }

            internal string Name { get; }
            internal string? Type { get; }
            internal uint? Index { get; }
            internal PowerPointLayoutRect Box { get; }
            internal string Anchor { get; }
            internal string? Size { get; }
            internal string? ListStyle { get; }
            internal string? Text { get; }
            internal string? Orientation { get; }
        }

        private static string ShapeXml(uint id, Placeholder ph, bool master) {
            var b = new StringBuilder();
            b.Append("<p:sp><p:nvSpPr><p:cNvPr id=\"").Append(id).Append("\" name=\"").Append(SecurityElement.Escape(ph.Name)).Append("\"/>");
            b.Append("<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph");
            if (ph.Type != null) b.Append(" type=\"").Append(ph.Type).Append('"');
            if (ph.Orientation != null) b.Append(" orient=\"").Append(ph.Orientation).Append('"');
            if (ph.Size != null) b.Append(" sz=\"").Append(ph.Size).Append('"');
            if (ph.Index.HasValue) b.Append(" idx=\"").Append(ph.Index.Value).Append('"');
            b.Append("/></p:nvPr></p:nvSpPr>");
            b.Append("<p:spPr><a:xfrm><a:off x=\"").Append(Pt(ph.Box.X)).Append("\" y=\"").Append(Pt(ph.Box.Y))
                .Append("\"/><a:ext cx=\"").Append(Pt(ph.Box.Width)).Append("\" cy=\"").Append(Pt(ph.Box.Height)).Append("\"/></a:xfrm>");
            if (master) b.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            b.Append("</p:spPr>");
            if (ph.Type == "pic") {
                b.Append("</p:sp>");
                return b.ToString();
            }
            b.Append("<p:txBody><a:bodyPr vert=\"").Append(ph.Orientation == "vert" ? "eaVert" : "horz")
                .Append("\" wrap=\"square\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"").Append(ph.Anchor).Append("\">");
            b.Append("<a:noAutofit/>");
            b.Append("</a:bodyPr>");
            // Placeholder list styles stay empty: binary PowerPoint materializes layout placeholders and has no
            // per-shape list style, so sizes live in the master text styles and on the slide runs instead.
            b.Append("<a:lstStyle/>");
            if (ph.Type == "sldNum") {
                b.Append("<a:p><a:fld id=\"{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}\" type=\"slidenum\"><a:rPr/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang=\"en-US\"/></a:p>");
            } else if (ph.Text != null) {
                b.Append("<a:p><a:r><a:rPr/><a:t>").Append(SecurityElement.Escape(ph.Text)).Append("</a:t></a:r><a:endParaRPr lang=\"en-US\"/></a:p>");
            } else {
                b.Append("<a:p><a:endParaRPr lang=\"en-US\"/></a:p>");
            }
            b.Append("</p:txBody></p:sp>");
            return b.ToString();
        }

        private static string FooterListStyle(PowerPointLayoutTheme theme, string align) {
            PowerPointTextStyleSpec footer = theme[PowerPointTextRole.Footer];
            return "<a:lstStyle><a:lvl1pPr algn=\"" + align + "\"><a:buNone/>" +
                   Run("+mn-lt", footer.Size, false, "<a:solidFill><a:srgbClr val=\"" + theme.TextMuted + "\"/></a:solidFill>") +
                   "</a:lvl1pPr></a:lstStyle>";
        }

        private static string StyleListStyle(PowerPointLayoutTheme theme, PowerPointTextRole role, bool bullets = false) {
            PowerPointTextStyleSpec s = theme[role];
            string font = s.HeadingFont ? "+mj-lt" : "+mn-lt";
            string color = s.Color == theme.Text ? Scheme("tx1") : "<a:solidFill><a:srgbClr val=\"" + s.Color + "\"/></a:solidFill>";
            return "<a:lstStyle><a:lvl1pPr marL=\"0\" indent=\"0\"><a:lnSpc><a:spcPct val=\"" + Pct(LineSpacingFraction(s.LineHeight)) + "\"/></a:lnSpc>" +
                   (bullets ? string.Empty : "<a:buNone/>") + Run(font, s.Size, s.Bold, color, s.Tracking) + "</a:lvl1pPr></a:lstStyle>";
        }

        private static IEnumerable<Placeholder> FooterPlaceholders(PowerPointLayoutTheme theme, PowerPointLayoutGrid grid, bool master) {
            yield return new Placeholder("Footer Placeholder", "ftr", master ? 3U : FooterIndex,
                new PowerPointLayoutRect(grid.X(1), grid.FooterY, grid.W(8), PowerPointLayoutGrid.FooterHeight), "b", "quarter",
                master ? null : FooterListStyle(theme, "l"));
            yield return new Placeholder("Slide Number Placeholder", "sldNum", master ? 4U : SlideNumberIndex,
                new PowerPointLayoutRect(grid.X(9), grid.FooterY, grid.W(4), PowerPointLayoutGrid.FooterHeight), "b", "quarter",
                master ? null : FooterListStyle(theme, "r"));
        }

        private static PowerPointLayoutRect TitleBox(PowerPointLayoutGrid grid) => grid.Box(1, 10, 2, 2);
        private static PowerPointLayoutRect ContentBox(PowerPointLayoutGrid grid) => grid.Box(1, 12, 4, 9);

        private static void WriteMaster(SlideMasterPart masterPart, PowerPointLayoutTheme theme, PowerPointLayoutGrid grid) {
            P.SlideMaster? master = masterPart.SlideMaster;
            if (master == null) return;
            var placeholders = new List<Placeholder> {
                new("Title Placeholder 1", "title", null, TitleBox(grid), text: "Click to edit Master title style"),
                new("Text Placeholder 2", "body", 1U, ContentBox(grid), text: "Click to edit Master text styles"),
            };
            placeholders.AddRange(FooterPlaceholders(theme, grid, master: true));
            string tree = TreeXml(placeholders, master: true);
            string background = "<p:bg xmlns:p=\"" + NsP + "\" xmlns:a=\"" + NsA + "\"><p:bgPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>";
            master.CommonSlideData = new P.CommonSlideData(new P.Background(background), new P.ShapeTree(tree));
            master.TextStyles = new P.TextStyles(TextStylesXml(theme));
            master.Save();
        }

        private static string TreeXml(IReadOnlyList<Placeholder> placeholders, bool master) {
            var b = new StringBuilder();
            b.Append("<p:spTree xmlns:p=\"").Append(NsP).Append("\" xmlns:a=\"").Append(NsA).Append("\">");
            b.Append("<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
            b.Append("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");
            uint id = 2;
            foreach (Placeholder ph in placeholders) b.Append(ShapeXml(id++, ph, master));
            b.Append("</p:spTree>");
            return b.ToString();
        }

        private static void WriteLayout(SlideLayoutPart layoutPart, PowerPointLayoutTheme theme, PowerPointLayoutGrid grid) {
            P.SlideLayout? layout = layoutPart.SlideLayout;
            if (layout == null) return;
            string name = layout.CommonSlideData?.Name?.Value ?? string.Empty;
            List<Placeholder>? placeholders = LayoutPlaceholders(name, theme, grid);
            if (placeholders == null) return;
            // Footer and slide-number placeholders live on the master only: binary PowerPoint materializes every layout
            // placeholder onto its slides, so layout footers would add empty shapes to every slide saved as .ppt.
            layout.CommonSlideData = new P.CommonSlideData(new P.ShapeTree(TreeXml(placeholders, master: false))) { Name = name };
            layout.Save();
        }

        private static List<Placeholder>? LayoutPlaceholders(string name, PowerPointLayoutTheme theme, PowerPointLayoutGrid grid) {
            Placeholder Title(PowerPointLayoutRect? box = null) => new("Title 1", "title", 0U, box ?? TitleBox(grid), text: "Click to edit Master title style");
            Placeholder Body(string label, uint index, PowerPointLayoutRect box, string? listStyle = null) =>
                new(label, "body", index, box, listStyle: listStyle, text: "Click to edit Master text styles");
            switch (name) {
                case "Title Slide":
                    return new List<Placeholder> {
                        new("Title 1", "ctrTitle", 0U, grid.Box(1, 9, 3, 6), "b", listStyle: StyleListStyle(theme, PowerPointTextRole.Display),
                            text: "Click to edit Master title style"),
                        new("Subtitle 2", "subTitle", 1U, grid.Box(1, 9, 9, 2), listStyle: StyleListStyle(theme, PowerPointTextRole.Subtitle),
                            text: "Click to edit Master subtitle style"),
                    };
                case "Title and Content":
                    return new List<Placeholder> { Title(), Body("Content Placeholder 2", 1U, ContentBox(grid)) };
                case "Section Header":
                    return new List<Placeholder> {
                        new("Title 1", "title", 0U, grid.Box(1, 8, 4, 5), "b", listStyle: StyleListStyle(theme, PowerPointTextRole.SectionTitle),
                            text: "Click to edit Master title style"),
                        new("Text Placeholder 2", "body", 1U, grid.Box(1, 8, 9, 2), listStyle: StyleListStyle(theme, PowerPointTextRole.Subtitle),
                            text: "Click to edit Master text styles"),
                    };
                case "Two Content":
                    return new List<Placeholder> {
                        Title(),
                        Body("Content Placeholder 2", 1U, grid.Box(1, 6, 4, 9)),
                        Body("Content Placeholder 3", 2U, grid.Box(7, 6, 4, 9)),
                    };
                case "Comparison":
                    string heading = StyleListStyle(theme, PowerPointTextRole.Heading);
                    return new List<Placeholder> {
                        Title(),
                        Body("Text Placeholder 2", 1U, grid.Box(1, 6, 4, 1), heading),
                        Body("Content Placeholder 3", 2U, grid.Box(1, 6, 5, 8)),
                        Body("Text Placeholder 4", 3U, grid.Box(7, 6, 4, 1), heading),
                        Body("Content Placeholder 5", 4U, grid.Box(7, 6, 5, 8)),
                    };
                case "Title Only":
                    return new List<Placeholder> { Title() };
                case "Blank":
                    return new List<Placeholder>();
                case "Picture with Caption":
                    return new List<Placeholder> {
                        Title(grid.Box(1, 5, 2, 2)),
                        new("Picture Placeholder 2", "pic", 1U, grid.Box(6, 7, 1, 12)),
                        Body("Text Placeholder 3", 2U, grid.Box(1, 5, 4, 7), StyleListStyle(theme, PowerPointTextRole.BodySmall)),
                    };
                case "Title and Vertical Text":
                    return new List<Placeholder> {
                        Title(),
                        new("Vertical Text Placeholder 2", "body", 1U, ContentBox(grid), orientation: "vert", text: "Click to edit Master text styles"),
                    };
                case "Vertical Title and Text":
                    return new List<Placeholder> {
                        new("Vertical Title 1", "title", null, grid.Box(10, 3, 1, 12), orientation: "vert", text: "Click to edit Master title style"),
                        new("Vertical Text Placeholder 2", "body", 1U, grid.Box(1, 9, 1, 12), orientation: "vert", text: "Click to edit Master text styles"),
                    };
                case "Two Content with Caption":
                    return new List<Placeholder> {
                        Title(),
                        Body("Content Placeholder 2", 1U, grid.Box(1, 6, 4, 7)),
                        Body("Content Placeholder 3", 2U, grid.Box(7, 6, 4, 7)),
                        Body("Caption Placeholder 4", 3U, grid.Box(1, 12, 11, 2), StyleListStyle(theme, PowerPointTextRole.Caption)),
                    };
                default:
                    return null;
            }
        }
    }
}
