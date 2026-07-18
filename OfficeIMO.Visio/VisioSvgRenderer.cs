using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        private const string SvgNamespace = "http://www.w3.org/2000/svg";

        public static string Render(VisioPage page, VisioSvgSaveOptions options) {
            if (options.PixelsPerInch <= 0D || double.IsNaN(options.PixelsPerInch) || double.IsInfinity(options.PixelsPerInch)) {
                throw new ArgumentOutOfRangeException(nameof(options), "PixelsPerInch must be a finite positive number.");
            }

            double scale = options.PixelsPerInch;
            double logicalWidth = Math.Max(page.Width, 0.01D) * scale;
            double logicalHeight = Math.Max(page.Height, 0.01D) * scale;
            double surfaceWidth = Math.Ceiling(logicalWidth);
            double surfaceHeight = Math.Ceiling(logicalHeight);

            StringBuilder builder = new();
            XmlWriterSettings settings = new() {
                OmitXmlDeclaration = !options.IncludeXmlDeclaration,
                Indent = true
            };

            using (XmlWriter writer = XmlWriter.Create(new Utf8StringWriter(builder), settings)) {
                writer.WriteStartDocument();
                writer.WriteStartElement("svg", SvgNamespace);
                writer.WriteNumberAttribute("width", surfaceWidth);
                writer.WriteNumberAttribute("height", surfaceHeight);
                writer.WriteViewBoxAttribute(0D, 0D, logicalWidth, logicalHeight);
                writer.WriteAttributeString("role", "img");
                writer.WriteAttributeString("aria-label", string.IsNullOrWhiteSpace(page.Name) ? "OfficeIMO Visio page" : page.Name);
                WriteEmbeddedFonts(writer, options.Fonts);

                if (options.BackgroundColor.HasValue && options.BackgroundColor.Value.A > 0) {
                    writer.WriteStartElement("rect", SvgNamespace);
                    writer.WriteNumberAttribute("x", 0D);
                    writer.WriteNumberAttribute("y", 0D);
                    writer.WriteNumberAttribute("width", logicalWidth);
                    writer.WriteNumberAttribute("height", logicalHeight);
                    OfficeSvgFormatting.WriteColorAttribute(writer, "fill", options.BackgroundColor.Value);
                    writer.WriteEndElement();
                }

                writer.WriteStartElement("g", SvgNamespace);
                writer.WriteAttributeString("data-officeimo-visio-page", page.Name);

                foreach (VisioShape shape in page.Shapes) {
                    WriteShape(writer, page, shape, options, scale);
                }

                VisioRenderLabelLayout? labelLayout = options.ResolveConnectorLabelOverlaps
                    ? VisioRenderLabelLayout.Create(page)
                    : null;
                foreach (VisioConnector connector in page.Connectors) {
                    WriteConnector(writer, page, connector, options, scale, labelLayout);
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }

            return builder.ToString();
        }

        private static void WriteEmbeddedFonts(XmlWriter writer, OfficeFontFaceCollection fonts) {
            if (fonts == null || fonts.Faces.Count == 0) return;
            var css = new StringBuilder();
            foreach (OfficeFontFace face in fonts.Faces) {
                css.Append("@font-face{font-family:\"")
                    .Append(EscapeCssString(face.FamilyName))
                    .Append("\";src:url(data:font/ttf;base64,")
                    .Append(Convert.ToBase64String(face.Data))
                    .Append(") format(\"truetype\");font-weight:")
                    .Append((face.Style & OfficeFontStyle.Bold) == OfficeFontStyle.Bold ? "700" : "400")
                    .Append(";font-style:")
                    .Append((face.Style & OfficeFontStyle.Italic) == OfficeFontStyle.Italic ? "italic" : "normal")
                    .Append(";}");
            }

            writer.WriteStartElement("defs", SvgNamespace);
            writer.WriteStartElement("style", SvgNamespace);
            writer.WriteAttributeString("type", "text/css");
            writer.WriteString(css.ToString());
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static string EscapeCssString(string value) =>
            value.Replace("\\", "\\\\").Replace("\"", "\\\"");

        private sealed class Utf8StringWriter : StringWriter {
            internal Utf8StringWriter(StringBuilder builder) : base(builder, CultureInfo.InvariantCulture) {
            }

            public override Encoding Encoding => Encoding.UTF8;
        }

        private static void WriteShape(XmlWriter writer, VisioPage page, VisioShape shape, VisioSvgSaveOptions options, double scale) {
            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-visio-shape-id", shape.Id);
            if (!string.IsNullOrWhiteSpace(shape.NameU)) {
                writer.WriteAttributeString("data-visio-nameu", shape.NameU);
            }

            WriteShapeGeometry(writer, page, shape, scale);

            if (options.RenderStencilArtwork) {
                if (!WritePackagePreviewArtwork(writer, page, shape, options, scale)) {
                    WriteStencilArtwork(writer, page, shape, scale);
                }
            }

            if (options.RenderText && !string.IsNullOrEmpty(shape.Text)) {
                WriteShapeText(writer, page, shape, scale);
            }

            foreach (VisioShape child in shape.Children) {
                WriteShape(writer, page, child, options, scale);
            }

            writer.WriteEndElement();
        }

    }
}
