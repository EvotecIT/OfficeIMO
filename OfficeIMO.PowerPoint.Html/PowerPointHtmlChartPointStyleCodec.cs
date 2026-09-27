using System;
using System.Globalization;
using System.Text;
using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

internal static class PowerPointHtmlChartPointStyleCodec {
    internal static void Write(StringBuilder output, OfficeChartPointStyle? style) {
        if (style == null) return;
        void Attribute(string key, string value) => output.Append(" data-officeimo-point-").Append(key)
            .Append("=\"").Append(OfficeHtmlText.EscapeAttribute(value)).Append('"');
        if (style.FillColor.HasValue) Attribute("fill", style.FillColor.Value.ToString());
        if (style.NoFill) Attribute("no-fill", "true");
        if (style.Hatch.HasValue) {
            Attribute("hatch", style.Hatch.Value.ToString());
            Attribute("hatch-color", style.HatchColor!.Value.ToString());
        }
        if (style.OutlineColor.HasValue) Attribute("outline-color", style.OutlineColor.Value.ToString());
        if (style.OutlineWidth.HasValue) Attribute("outline-width", style.OutlineWidth.Value.ToString("G17", CultureInfo.InvariantCulture));
        if (style.ShowOutline.HasValue) Attribute("show-outline", style.ShowOutline.Value ? "true" : "false");
    }

    internal static bool TryRead(IElement cell, out OfficeChartPointStyle? style) {
        style = null;
        string? Get(string key) => cell.GetAttribute("data-officeimo-point-" + key);
        string? rawFill = Get("fill"), rawNoFill = Get("no-fill"), rawHatch = Get("hatch"),
            rawHatchColor = Get("hatch-color"), rawOutline = Get("outline-color"),
            rawWidth = Get("outline-width"), rawShow = Get("show-outline");
        if (rawFill == null && rawNoFill == null && rawHatch == null && rawHatchColor == null &&
            rawOutline == null && rawWidth == null && rawShow == null) return true;
        bool Color(string? raw, out OfficeColor? color) {
            color = null;
            if (raw == null) return true;
            if (!OfficeColor.TryParse(raw, out OfficeColor parsed)) return false;
            color = parsed;
            return true;
        }
        if (!Color(rawFill, out OfficeColor? fill) || !Color(rawHatchColor, out OfficeColor? hatchColor) ||
            !Color(rawOutline, out OfficeColor? outline)) return false;
        bool noFill = false;
        bool? show = null;
        if (rawNoFill != null && !bool.TryParse(rawNoFill, out noFill)) return false;
        if (rawShow != null) {
            if (!bool.TryParse(rawShow, out bool value)) return false;
            show = value;
        }
        OfficeChartHatchPattern? hatch = null;
        if (rawHatch != null) {
            if (!Enum.TryParse(rawHatch, out OfficeChartHatchPattern parsed) ||
                !Enum.IsDefined(typeof(OfficeChartHatchPattern), parsed)) return false;
            hatch = parsed;
        }
        double? width = null;
        if (rawWidth != null) {
            if (!double.TryParse(rawWidth, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) return false;
            width = parsed;
        }
        try {
            style = new OfficeChartPointStyle(fill, noFill, hatch, hatchColor, outline, width, show);
            return true;
        } catch (ArgumentException) { return false; }
    }
}
