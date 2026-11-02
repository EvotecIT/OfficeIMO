using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>One text run written by the layout system.</summary>
    internal sealed class PowerPointLayoutRun {
        internal PowerPointLayoutRun(string text, PowerPointTextStyleSpec style, double? size = null, string? color = null,
            bool? bold = null) {
            Text = style.Upper ? (text ?? string.Empty).ToUpperInvariant() : text ?? string.Empty;
            Heading = style.HeadingFont;
            Size = size ?? style.Size;
            Color = color ?? style.Color;
            Bold = bold ?? style.Bold;
            Tracking = style.Tracking;
        }

        internal string Text { get; }
        internal bool Heading { get; set; }
        internal double Size { get; }
        internal string Color { get; }
        internal OfficeColor? ExplicitColor { get; set; }
        internal double Alpha { get; set; } = 1D;
        internal bool Bold { get; }
        internal double Tracking { get; set; }
        internal string? Font { get; set; }
        internal bool Field { get; set; }
    }

    /// <summary>One paragraph written by the layout system.</summary>
    internal sealed class PowerPointLayoutParagraph {
        internal PowerPointLayoutParagraph(params PowerPointLayoutRun[] runs) {
            Runs = runs.ToList();
        }

        internal List<PowerPointLayoutRun> Runs { get; }
        internal PowerPointParagraph? NativeParagraph { get; set; }
        internal PowerPointLayoutNativeText? NativeBody { get; set; }
        internal string Align { get; set; } = "l";
        internal double LineHeight { get; set; } = 1.2D;
        internal double SpaceAfter { get; set; }
        internal double SpaceBefore { get; set; }
        internal int Level { get; set; } = 1;

        /// <summary>Bullet character; null for none.</summary>
        internal string? Bullet { get; set; }

        internal string? BulletColor { get; set; }
        internal double MarginLeft { get; set; }
        internal double Indent { get; set; }

        /// <summary>Auto-numbering scheme (arabicPeriod) with a start value; null for none.</summary>
        internal int? NumberStart { get; set; }
    }

    /// <summary>Builds the native shapes the layout system writes: text frames, rules, panels, markers.</summary>
    internal sealed class PowerPointLayoutWriter {
        private readonly PowerPointSlide _slide;
        private readonly PowerPointLayoutTheme _theme;

        internal PowerPointLayoutWriter(PowerPointSlide slide, PowerPointLayoutTheme theme) {
            _slide = slide;
            _theme = theme;
        }

        internal PowerPointSlide Slide => _slide;

        private static string E(double points) => ((long)Math.Round(points * 12700D)).ToString(CultureInfo.InvariantCulture);
        private static string I(double value) => ((int)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
        private static string X(string text) => SecurityElement.Escape(text) ?? string.Empty;

        /// <summary>A colour as a theme reference when it is a theme slot, otherwise a literal.</summary>
        internal string ColorXml(string hex, double alpha = 1D) {
            if (hex.Length == 8 && OfficeColor.TryParse("#" + hex, out OfficeColor literal)) {
                alpha *= literal.A / 255D;
                hex = literal.ToRgbHex();
            }
            string? scheme = null;
            if (string.Equals(hex, _theme.Text, StringComparison.OrdinalIgnoreCase)) scheme = "tx1";
            else if (string.Equals(hex, _theme.Background, StringComparison.OrdinalIgnoreCase)) scheme = "bg1";
            else if (string.Equals(hex, _theme.Accent1, StringComparison.OrdinalIgnoreCase)) scheme = "accent1";
            else if (string.Equals(hex, _theme.Accent2, StringComparison.OrdinalIgnoreCase)) scheme = "accent2";
            else if (string.Equals(hex, _theme.Accent3, StringComparison.OrdinalIgnoreCase)) scheme = "accent3";
            string a = alpha < 0.999D ? "<a:alpha val=\"" + I(alpha * 100000D) + "\"/>" : string.Empty;
            return scheme != null
                ? "<a:schemeClr val=\"" + scheme + "\">" + a + "</a:schemeClr>"
                : "<a:srgbClr val=\"" + hex + "\">" + a + "</a:srgbClr>";
        }

        private string ExplicitColorXml(OfficeColor color, double alpha) {
            if (color.ThemeSlot is string slot) {
                string slotOpacity = alpha < 0.999D ? "<a:alpha val=\"" + I(alpha * 100000D) + "\"/>" : string.Empty;
                return "<a:schemeClr val=\"" + slot + "\">" + slotOpacity + "</a:schemeClr>";
            }
            string rgb;
            if (color.IsThemeRole) {
                rgb = color.Role switch {
                    OfficeColorRole.Success => _theme.Success,
                    OfficeColorRole.Warning => _theme.Warning,
                    OfficeColorRole.Danger => _theme.Danger,
                    _ => throw new ArgumentException("The colour role has no supported native slot.", nameof(color))
                };
            } else {
                rgb = color.ToRgbHex();
                alpha *= color.A / 255D;
            }
            string opacity = alpha < 0.999D ? "<a:alpha val=\"" + I(alpha * 100000D) + "\"/>" : string.Empty;
            return "<a:srgbClr val=\"" + rgb + "\">" + opacity + "</a:srgbClr>";
        }

        private string Fill(string? hex, double alpha = 1D, OfficeColor? explicitColor = null) =>
            hex == null && !explicitColor.HasValue ? "<a:noFill/>" : "<a:solidFill>" +
                (explicitColor.HasValue ? ExplicitColorXml(explicitColor.Value, alpha) : ColorXml(hex!, alpha)) + "</a:solidFill>";

        private static string Header(string name, string? description, string? placeholder, bool decorative, bool textBox) {
            var b = new StringBuilder();
            b.Append("<p:sp xmlns:p=\"").Append(PowerPointLayoutMasterWriter.NsP).Append("\" xmlns:a=\"").Append(PowerPointLayoutMasterWriter.NsA).Append("\">");
            b.Append("<p:nvSpPr><p:cNvPr id=\"0\" name=\"").Append(X(name)).Append('"');
            if (!string.IsNullOrWhiteSpace(description)) b.Append(" descr=\"").Append(X(description!)).Append('"');
            b.Append('>');
            if (decorative) {
                b.Append("<a:extLst><a:ext uri=\"{C183D7F6-B498-43B3-948B-1728B52AA6E4}\"><adec:decorative xmlns:adec=\"http://schemas.microsoft.com/office/drawing/2017/decorative\" val=\"1\"/></a:ext></a:extLst>");
            }
            b.Append("</p:cNvPr>");
            b.Append(placeholder != null ? "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>" : textBox ? "<p:cNvSpPr txBox=\"1\"/>" : "<p:cNvSpPr/>");
            b.Append("<p:nvPr>").Append(placeholder ?? string.Empty).Append("</p:nvPr></p:nvSpPr>");
            return b.ToString();
        }

        private static string Xfrm(PowerPointLayoutRect r, double rotation = 0D) =>
            "<a:xfrm" + (Math.Abs(rotation) > 0.01D ? " rot=\"" + I(rotation * 60000D) + "\"" : string.Empty) + "><a:off x=\"" + E(r.X) + "\" y=\"" + E(r.Y) +
            "\"/><a:ext cx=\"" + E(Math.Max(0.1D, r.Width)) + "\" cy=\"" + E(Math.Max(0.1D, r.Height)) + "\"/></a:xfrm>";

        private string RunXml(PowerPointLayoutRun run) {
            string font = run.Font ?? (run.Heading ? "+mj-lt" : "+mn-lt");
            var b = new StringBuilder();
            b.Append("<a:rPr lang=\"en-US\" sz=\"").Append(I(run.Size * 100D)).Append("\" b=\"").Append(run.Bold ? '1' : '0').Append('"');
            if (Math.Abs(run.Tracking) > 0.0001D) b.Append(" spc=\"").Append(I(run.Tracking * run.Size * 100D)).Append('"');
            b.Append('>');
            b.Append("<a:solidFill>").Append(run.ExplicitColor.HasValue ? ExplicitColorXml(run.ExplicitColor.Value, run.Alpha) : ColorXml(run.Color, run.Alpha)).Append("</a:solidFill>");
            if (font.StartsWith("+", StringComparison.Ordinal)) {
                b.Append("<a:latin typeface=\"").Append(font).Append("\"/>");
            } else {
                var latin = new A.LatinFont();
                PowerPointFontClassification.Apply(latin, font);
                b.Append("<a:latin typeface=\"").Append(X(font)).Append('"');
                if (latin.Panose != null) b.Append(" panose=\"").Append(latin.Panose.Value).Append('"');
                if (latin.PitchFamily != null) b.Append(" pitchFamily=\"").Append(latin.PitchFamily.Value).Append('"');
                b.Append(" charset=\"0\"/>");
            }
            b.Append("</a:rPr>");
            string props = b.ToString();
            if (run.Field) {
                return "<a:fld id=\"{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}\" type=\"slidenum\">" + props + "<a:t>" + X(run.Text) + "</a:t></a:fld>";
            }

            // Explicit line breaks inside a run become <a:br/>.
            string[] parts = run.Text.Split('\n');
            var r = new StringBuilder();
            for (int i = 0; i < parts.Length; i++) {
                if (i > 0) r.Append("<a:br>").Append(props).Append("</a:br>");
                r.Append("<a:r>").Append(props).Append("<a:t>").Append(X(parts[i])).Append("</a:t></a:r>");
            }

            return r.ToString();
        }

        private string ParagraphXml(PowerPointLayoutParagraph p) {
            var b = new StringBuilder("<a:p><a:pPr");
            if (p.Level > 1) b.Append(" lvl=\"").Append(p.Level - 1).Append('"');
            b.Append(" marL=\"").Append(E(p.MarginLeft)).Append("\" indent=\"").Append(E(p.Indent)).Append("\" algn=\"").Append(p.Align).Append("\">");
            double lineSize = p.Runs.Count == 0 ? 12D : p.Runs.Max(run => run.Size);
            if (p.NativeParagraph != null) lineSize = PowerPointLayoutNativeTextMeasure.MaximumSize(p.NativeParagraph, lineSize);
            b.Append("<a:lnSpc><a:spcPts val=\"").Append(PowerPointLayoutMasterWriter.LineSpacingPoints(lineSize, p.LineHeight)).Append("\"/></a:lnSpc>");
            b.Append("<a:spcBef><a:spcPts val=\"").Append(I(p.SpaceBefore * 100D)).Append("\"/></a:spcBef>");
            b.Append("<a:spcAft><a:spcPts val=\"").Append(I(p.SpaceAfter * 100D)).Append("\"/></a:spcAft>");
            if (p.NumberStart.HasValue) {
                b.Append("<a:buClr>").Append(ColorXml(p.BulletColor ?? _theme.Accent1)).Append("</a:buClr><a:buSzPct val=\"100000\"/>");
                b.Append("<a:buFont typeface=\"+mj-lt\"/><a:buAutoNum type=\"arabicPeriod\" startAt=\"").Append(p.NumberStart.Value).Append("\"/>");
            } else if (p.Bullet != null) {
                b.Append("<a:buClr>").Append(ColorXml(p.BulletColor ?? _theme.Accent1)).Append("</a:buClr><a:buSzPct val=\"100000\"/>");
                b.Append("<a:buFont typeface=\"Arial\" panose=\"020B0604020202020204\" pitchFamily=\"34\" charset=\"0\"/><a:buChar char=\"").Append(X(p.Bullet)).Append("\"/>");
            } else {
                b.Append("<a:buNone/>");
            }
            b.Append("</a:pPr>");
            foreach (PowerPointLayoutRun run in p.Runs) b.Append(RunXml(run));
            PowerPointLayoutRun? last = p.Runs.LastOrDefault();
            b.Append("<a:endParaRPr lang=\"en-US\" sz=\"").Append(I((last?.Size ?? 12D) * 100D)).Append("\"/>");
            b.Append("</a:p>");
            return b.ToString();
        }

        /// <summary>
        ///     Adds a text frame. Frames have zero insets so text edges sit on the grid, never autofit (the layout
        ///     measured the text) and wrap within their width.
        /// </summary>
        internal PowerPointTextBox Text(string name, PowerPointLayoutRect rect, IEnumerable<PowerPointLayoutParagraph> paragraphs,
            string anchor = "t", string? placeholder = null, string? description = null, bool decorative = false,
            string? fill = null, double inset = 0D, bool wrap = true, PowerPointLayoutNativeText? nativeBody = null) {
            var b = new StringBuilder(Header(name, description, placeholder, decorative, textBox: true));
            b.Append("<p:spPr>").Append(Xfrm(rect)).Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            b.Append(Fill(fill)).Append("<a:ln><a:noFill/></a:ln></p:spPr>");
            string ins = E(inset);
            b.Append("<p:txBody><a:bodyPr wrap=\"").Append(wrap ? "square" : "none").Append("\" lIns=\"").Append(ins).Append("\" tIns=\"").Append(ins)
                .Append("\" rIns=\"").Append(ins).Append("\" bIns=\"").Append(ins).Append("\" anchor=\"").Append(anchor).Append("\"><a:noAutofit/></a:bodyPr><a:lstStyle/>");
            List<PowerPointLayoutParagraph> list = paragraphs.ToList();
            if (list.Count == 0) b.Append("<a:p><a:endParaRPr lang=\"en-US\"/></a:p>");
            foreach (PowerPointLayoutParagraph paragraph in list) b.Append(ParagraphXml(paragraph));
            b.Append("</p:txBody></p:sp>");
            var generated = (PowerPointTextBox)_slide.AppendLayoutShape(new Shape(b.ToString()), textBox: true);
            return PowerPointLayoutNativeText.Bind(_slide, generated, list, nativeBody);
        }

        internal PowerPointTextBox Text(string name, PowerPointLayoutRect rect, PowerPointLayoutParagraph paragraph,
            string anchor = "t", string? placeholder = null, string? description = null, bool decorative = false) =>
            Text(name, rect, new[] { paragraph }, anchor, placeholder, description, decorative);

        /// <summary>A filled rectangle (rules, panels, scrims). <paramref name="outline" /> draws a hairline frame.</summary>
        internal PowerPointAutoShape Rect(string name, PowerPointLayoutRect rect, string? fill, double fillAlpha = 1D,
            string? outline = null, double outlineWidth = 1D, bool decorative = true, string? gradient = null, OfficeColor? explicitFill = null) {
            var b = new StringBuilder(Header(name, null, null, decorative, textBox: false));
            b.Append("<p:spPr>").Append(Xfrm(rect)).Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            b.Append(gradient ?? Fill(fill, fillAlpha, explicitFill));
            b.Append(outline == null ? "<a:ln><a:noFill/></a:ln>" : "<a:ln w=\"" + E(outlineWidth) + "\">" + Fill(outline) + "</a:ln>");
            b.Append("</p:spPr></p:sp>");
            return (PowerPointAutoShape)_slide.AppendLayoutShape(new Shape(b.ToString()), textBox: false);
        }

        /// <summary>A linear gradient of one colour; each stop is (position 0..1, alpha 0..1). Angle in degrees (90 = top to bottom).</summary>
        internal string Gradient(string hex, int angle, params (double Position, double Alpha)[] stops) {
            var b = new StringBuilder("<a:gradFill rotWithShape=\"1\"><a:gsLst>");
            foreach ((double position, double alpha) in stops) {
                b.Append("<a:gs pos=\"").Append(I(position * 100000D)).Append("\">").Append(ColorXml(hex, Math.Max(0D, alpha))).Append("</a:gs>");
            }
            b.Append("</a:gsLst><a:lin ang=\"").Append(I(angle * 60000D)).Append("\" scaled=\"0\"/></a:gradFill>");
            return b.ToString();
        }

        /// <summary>A circle marker with centred text (process steps, map pins).</summary>
        internal PowerPointShape Circle(string name, PowerPointLayoutRect rect, string? fill, string? stroke, double strokeWidth,
            PowerPointLayoutParagraph? label, double fillAlpha = 1D, bool decorative = false, bool shape = false) {
            var b = new StringBuilder(Header(name, null, null, decorative, textBox: false));
            b.Append("<p:spPr>").Append(Xfrm(rect)).Append("<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>");
            b.Append(Fill(fill, fillAlpha));
            b.Append(stroke == null ? "<a:ln><a:noFill/></a:ln>" : "<a:ln w=\"" + E(strokeWidth) + "\">" + Fill(stroke) + "</a:ln>");
            b.Append("</p:spPr>");
            b.Append("<p:txBody><a:bodyPr wrap=\"none\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"ctr\" anchorCtr=\"0\"><a:noAutofit/></a:bodyPr><a:lstStyle/>");
            if (label != null) {
                label.Align = "ctr";
                b.Append(ParagraphXml(label));
            } else {
                b.Append("<a:p><a:endParaRPr lang=\"en-US\"/></a:p>");
            }
            b.Append("</p:txBody></p:sp>");
            return _slide.AppendLayoutShape(new Shape(b.ToString()), textBox: !shape);
        }

        /// <summary>A straight connector line, optionally with an arrowhead at the end point.</summary>
        internal PowerPointAutoShape Line(string name, double x1, double y1, double x2, double y2, string color, double width, bool arrow,
            bool dotted = false, double dotPitch = 12D, bool decorative = false) {
            var b = new StringBuilder(Header(name, null, null, decorative, textBox: false));
            double x = Math.Min(x1, x2), y = Math.Min(y1, y2);
            string flip = (x2 < x1 ? " flipH=\"1\"" : string.Empty) + (y2 < y1 ? " flipV=\"1\"" : string.Empty);
            b.Append("<p:spPr><a:xfrm").Append(flip).Append("><a:off x=\"").Append(E(x)).Append("\" y=\"").Append(E(y))
                .Append("\"/><a:ext cx=\"").Append(E(Math.Abs(x2 - x1))).Append("\" cy=\"").Append(E(Math.Abs(y2 - y1))).Append("\"/></a:xfrm>");
            b.Append("<a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom><a:noFill/>");
            b.Append("<a:ln w=\"").Append(E(width)).Append(dotted ? "\" cap=\"rnd\">" : "\">").Append(Fill(color));
            if (dotted) {
                // A dash as long as the line is wide, then a gap to the next dot: a row of round dots at dotPitch.
                b.Append("<a:custDash><a:ds d=\"100000\" sp=\"").Append(I((dotPitch / width - 2D) * 100000D)).Append("\"/></a:custDash>");
            }
            if (arrow) b.Append("<a:tailEnd type=\"triangle\" w=\"med\" len=\"med\"/>");
            b.Append("</a:ln></p:spPr></p:sp>");
            return (PowerPointAutoShape)_slide.AppendLayoutShape(new Shape(b.ToString()), textBox: false);
        }

        /// <summary>Sets a solid slide background.</summary>
        internal void Background(string hex) {
            Slide.BackgroundColor = hex;
        }
    }
}
