using System.Text;
using System.Xml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        /// <summary>Writes the shape's cached FillGradient as a user-space SVG gradient and returns its paint reference.</summary>
        private static string? WriteFillGradient(XmlWriter writer, VisioPage page, VisioShape shape, VisioRenderProjection projection) {
            if (shape.FillPattern == 0 || shape.FillColor.A == 0 ||
                !VisioFillGradient.TryResolve(shape, out VisioFillGradient? gradient, out _)) return null;
            VisioNativeShapeTransform transform = VisioNativeShapeTransform.Create(shape);
            OfficePoint Map(double x, double y) {
                OfficePoint point = transform.PagePoint(x, y);
                (double sx, double sy) = ToSvg(page, point.X, point.Y, projection);
                return new OfficePoint(sx, sy);
            }
            if (!gradient!.TryPlace(shape, Map, out OfficePoint start, out OfficePoint end, out OfficeTransform frame)) return null;
            string id = "visio-fill-" + page.Id.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + SvgId(shape.Id);
            writer.WriteStartElement("defs", SvgNamespace);
            if (gradient.IsLinear) {
                writer.WriteStartElement("linearGradient", SvgNamespace);
                writer.WriteAttributeString("id", id);
                writer.WriteAttributeString("gradientUnits", "userSpaceOnUse");
                writer.WriteNumberAttribute("x1", start.X);
                writer.WriteNumberAttribute("y1", start.Y);
                writer.WriteNumberAttribute("x2", end.X);
                writer.WriteNumberAttribute("y2", end.Y);
            } else {
                writer.WriteStartElement("radialGradient", SvgNamespace);
                writer.WriteAttributeString("id", id);
                writer.WriteAttributeString("gradientUnits", "userSpaceOnUse");
                writer.WriteAttributeString("cx", "0.5");
                writer.WriteAttributeString("cy", "0.5");
                writer.WriteAttributeString("r", "0.5");
                writer.WriteAttributeString("gradientTransform", "matrix(" + Format(frame.M11) + " " + Format(frame.M12) + " " +
                    Format(frame.M21) + " " + Format(frame.M22) + " " + Format(frame.OffsetX) + " " + Format(frame.OffsetY) + ")");
            }
            foreach (OfficeGradientStop stop in gradient.Stops) {
                writer.WriteStartElement("stop", SvgNamespace);
                writer.WriteNumberAttribute("offset", stop.Offset);
                writer.WriteAttributeString("stop-color", OfficeSvgFormatting.ToCssColor(OfficeColor.FromRgb(stop.Color.R, stop.Color.G, stop.Color.B)));
                if (stop.Color.A < 255) writer.WriteNumberAttribute("stop-opacity", stop.Color.A / 255D);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            writer.WriteEndElement();
            return "url(#" + id + ")";
        }

        private static string SvgId(string value) {
            var builder = new StringBuilder(value.Length);
            foreach (char character in value) builder.Append(char.IsLetterOrDigit(character) || character == '-' ? character : '_');
            return builder.ToString();
        }
    }
}
