using System.Globalization;
using System.Security;
using System.Text;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

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

        private const int D = Ns.DrawingML;
        private const int P = Ns.PresentationML;
        private const string TableStylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";

        /// <param name="presentationPart">Presentation to restyle.</param>
        /// <param name="theme">Theme tokens to write.</param>
        /// <param name="selectedMaster">Single master to restyle, or none for every master and the presentation defaults.</param>
        /// <param name="keepMasterContent">
        ///     True for a deck OfficeIMO did not author: only the theme part and text styles change, so shape trees,
        ///     backgrounds and placeholder links survive.
        /// </param>
        internal static void Apply(PowerPointPartRef presentationPart, PowerPointLayoutTheme theme,
            PowerPointPartRef selectedMaster = default, bool keepMasterContent = false) {
            OfficeXmlNode presentation = presentationPart.Root;
            if (presentation.IsNone) return;
            OfficeXmlNode size = presentation.Child(P, "sldSz");
            double width = (size.Int32Attribute("cx") ?? 12192000) / 12700D;
            double height = (size.Int32Attribute("cy") ?? 6858000) / 12700D;
            var grid = new PowerPointLayoutGrid(width, height, theme.MarginLeft, theme.MarginRight, theme.MarginTop);

            if (selectedMaster.IsNone)
                PowerPointDocument.ReplaceWithLayoutMarkup(presentation, DefaultTextStyleXml(theme));

            foreach (PowerPointPartRef masterPart in presentationPart.RelatedParts(PowerPointPartRef.SlideMasterRelationshipType).ToList()) {
                if (!selectedMaster.IsNone && !masterPart.Equals(selectedMaster)) continue;
                PowerPointPartRef themePart = masterPart.RelatedPart(PowerPointPartRef.ThemeRelationshipType);
                if (!themePart.IsNone) WriteTheme(themePart, theme);
                if (keepMasterContent) {
                    WriteTextStyles(masterPart, theme);
                    continue;
                }

                WriteMaster(masterPart, theme, grid);
                foreach (PowerPointPartRef layoutPart in masterPart.RelatedParts(PowerPointPartRef.SlideLayoutRelationshipType).ToList()) {
                    WriteLayout(layoutPart, theme, grid);
                }
            }

            if (selectedMaster.IsNone) WriteDefaultTableTextColor(presentationPart, theme);

            PowerPointDocument.SaveLayoutPart(presentationPart);
        }

        // The built-in table uses light accent tints even on a dark slide. Keep
        // its body text dark without assigning direct cell formatting, so an
        // explicitly selected table style and explicit cell colors still win.
        private static void WriteDefaultTableTextColor(PowerPointPartRef presentationPart, PowerPointLayoutTheme theme) {
            if (!theme.Theme.IsDark) return;
            PowerPointPartRef stylesPart = presentationPart.RelatedPart(TableStylesRelationshipType);
            if (stylesPart.IsNone) return;
            OfficeXmlNode style = default;
            foreach (OfficeXmlNode item in stylesPart.Root.Children(D, "tblStyle")) {
                if (item.Attribute("styleId") == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}") {
                    style = item;
                    break;
                }
            }
            OfficeXmlNode textStyle = style.Child(D, "wholeTbl").Child(D, "tcTxStyle");
            OfficeXmlNode inheritedColor = textStyle.Child(D, "schemeClr");
            if (inheritedColor.IsNone || inheritedColor.Attribute("val") != "dk1") return;
            inheritedColor.Remove();
            OfficeXmlNode color = textStyle.CreateElement(D, "srgbClr");
            color.SetAttribute("val", theme.Background);
            textStyle.Append(color);
            PowerPointDocument.SaveLayoutPart(stylesPart);
        }

        // ---------------------------------------------------------------- theme part

        internal static void WriteTheme(PowerPointPartRef themePart, PowerPointLayoutTheme theme) {
            OfficeXmlNode root = themePart.Root;
            if (root.IsNone) return;
            OfficeXmlNode elements = root.Child(D, "themeElements");
            OfficeXmlNode colors = elements.Child(D, "clrScheme");
            if (colors.Exists) {
                SetSchemeColor(colors.Child(D, "dk1"), theme.Text);
                SetSchemeColor(colors.Child(D, "lt1"), theme.Background);
                SetSchemeColor(colors.Child(D, "dk2"), theme.TextMuted);
                SetSchemeColor(colors.Child(D, "lt2"), theme.Surface);
                SetSchemeColor(colors.Child(D, "accent1"), theme.Accent1);
                SetSchemeColor(colors.Child(D, "accent2"), theme.Accent2);
                SetSchemeColor(colors.Child(D, "accent3"), theme.Accent3);
                SetSchemeColor(colors.Child(D, "accent4"), theme.Accent4);
                SetSchemeColor(colors.Child(D, "accent5"), theme.Accent5);
                SetSchemeColor(colors.Child(D, "accent6"), theme.Accent6);
                SetSchemeColor(colors.Child(D, "hlink"), theme.Accent1);
                SetSchemeColor(colors.Child(D, "folHlink"), theme.Accent5);
            }

            OfficeXmlNode fonts = elements.Child(D, "fontScheme");
            if (fonts.Exists) {
                SetSchemeFont(fonts.Child(D, "majorFont"), theme.HeadingFont);
                SetSchemeFont(fonts.Child(D, "minorFont"), theme.BodyFont);
            }

            PowerPointDocument.SaveLayoutPart(themePart);
        }

        private static void SetSchemeColor(OfficeXmlNode slot, string hex) {
            if (slot.IsNone) return;
            // Keep a system colour (windowText, window) when it already resolves to the requested value.
            OfficeXmlNode system = slot.Child(D, "sysClr");
            if (system.Exists && string.Equals(system.Attribute("lastClr"), hex, StringComparison.OrdinalIgnoreCase)) return;
            for (OfficeXmlNode child = slot.FirstChild; child.Exists; child = slot.FirstChild) child.Remove();
            OfficeXmlNode color = slot.CreateElement(D, "srgbClr");
            color.SetAttribute("val", hex);
            slot.Append(color);
        }

        private static void SetSchemeFont(OfficeXmlNode font, string typeface) {
            if (font.IsNone) return;
            PowerPointFontClassification.Apply(font.GetOrPrependChild(D, "latin"), typeface);
        }

        // ---------------------------------------------------------------- text styles

        private static string Pt(double points) => ((long)Math.Round(points * 12700D)).ToString(CultureInfo.InvariantCulture);
        // Paragraph spacing in whole 1/8 points (the binary PowerPoint field) expressed in 1/100 points.
        private static string Spc(double points) => ((int)Math.Round(points * 4D) * 25).ToString(CultureInfo.InvariantCulture);
        private static string Sz(double points) => ((int)Math.Round(points * 100D)).ToString(CultureInfo.InvariantCulture);
        // Absolute leading keeps the design's em-based measurement independent of each viewer's
        // natural font line height. Quarter-point spacing also survives binary PowerPoint storage.
        internal static string LineSpacingPoints(double size, double lineHeight) => Spc(size * lineHeight);

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
            b.Append("<a:lnSpc><a:spcPts val=\"").Append(LineSpacingPoints(title.Size, title.LineHeight)).Append("\"/></a:lnSpc>");
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
                b.Append("<a:lnSpc><a:spcPts val=\"").Append(LineSpacingPoints(size, body.LineHeight)).Append("\"/></a:lnSpc>");
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
            return "<a:lstStyle><a:lvl1pPr marL=\"0\" indent=\"0\"><a:lnSpc><a:spcPts val=\"" + LineSpacingPoints(s.Size, s.LineHeight) + "\"/></a:lnSpc>" +
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

        private static void WriteMaster(PowerPointPartRef masterPart, PowerPointLayoutTheme theme, PowerPointLayoutGrid grid) {
            OfficeXmlNode master = masterPart.Root;
            if (master.IsNone) return;
            var placeholders = new List<Placeholder> {
                new("Title Placeholder 1", "title", null, TitleBox(grid), text: "Click to edit Master title style"),
                new("Text Placeholder 2", "body", 1U, ContentBox(grid), text: "Click to edit Master text styles"),
            };
            placeholders.AddRange(FooterPlaceholders(theme, grid, master: true));
            string tree = TreeXml(placeholders, master: true);
            string background = "<p:bg xmlns:p=\"" + NsP + "\" xmlns:a=\"" + NsA + "\"><p:bgPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>";
            OfficeXmlNode data = master.CreateElement(P, "cSld");
            PowerPointDocument.ReplaceWithLayoutMarkup(data, background);
            PowerPointDocument.ReplaceWithLayoutMarkup(data, tree);
            PowerPointLayoutNativeText.Replace(master, P, "cSld", data, PSlideMasterView.Order);
            WriteTextStyles(masterPart, theme);
        }

        private static void WriteTextStyles(PowerPointPartRef masterPart, PowerPointLayoutTheme theme) {
            OfficeXmlNode master = masterPart.Root;
            if (master.IsNone) return;
            PowerPointDocument.ReplaceWithLayoutMarkup(master, TextStylesXml(theme));
            PowerPointDocument.SaveLayoutPart(masterPart);
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

        private static void WriteLayout(PowerPointPartRef layoutPart, PowerPointLayoutTheme theme, PowerPointLayoutGrid grid) {
            OfficeXmlNode layout = layoutPart.Root;
            if (layout.IsNone) return;
            string name = layout.Child(P, "cSld").Attribute("name") ?? string.Empty;
            List<Placeholder>? placeholders = LayoutPlaceholders(name, theme, grid);
            if (placeholders == null) return;
            // Footer and slide-number placeholders live on the master only: binary PowerPoint materializes every layout
            // placeholder onto its slides, so layout footers would add empty shapes to every slide saved as .ppt.
            OfficeXmlNode data = layout.CreateElement(P, "cSld");
            PowerPointDocument.ReplaceWithLayoutMarkup(data, TreeXml(placeholders, master: false));
            data.SetAttribute("name", name);
            PowerPointLayoutNativeText.Replace(layout, P, "cSld", data, PSlideLayoutView.Order);
            PowerPointDocument.SaveLayoutPart(layoutPart);
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
