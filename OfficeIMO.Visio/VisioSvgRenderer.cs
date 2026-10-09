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

        public static string Render(VisioPage page, VisioSvgSaveOptions options, VisioRenderLayerVisibility layerVisibility) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (options.PixelsPerInch <= 0D || double.IsNaN(options.PixelsPerInch) || double.IsInfinity(options.PixelsPerInch)) {
                throw new ArgumentOutOfRangeException(nameof(options), "PixelsPerInch must be a finite positive number.");
            }

            VisioRenderProjection projection = VisioRenderProjection.Create(page, options.PixelsPerInch);
            double logicalWidth = projection.WidthInches * projection.PhysicalDensity;
            double logicalHeight = projection.HeightInches * projection.PhysicalDensity;
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
                WriteEmbeddedFonts(writer, options.Fonts, options.CancellationToken);

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

                int foreignImageIndex = 0;
                foreach (VisioPage contentPage in VisioBackgroundComposition.Resolve(page, options.CancellationToken, options.ImageDiagnostics, options.ImageDiagnosticSource)) {
                    bool background = !ReferenceEquals(contentPage, page);
                    if (background) {
                        writer.WriteStartElement("g", SvgNamespace);
                        writer.WriteAttributeString("data-officeimo-visio-background", contentPage.Name);
                    }
                    var contentProjection = VisioRenderProjection.CreateForContent(contentPage, page, options.PixelsPerInch);
                    var contentVisibility = background ? new VisioRenderLayerVisibility(contentPage, options.LayerMode) : layerVisibility;
                    var textStyles = new VisioNativeTextStyleResolver(contentPage.OwnerDocument, options.CancellationToken, options.ImageDiagnostics, options.ImageDiagnosticSource);
                    foreach (VisioShape shape in contentPage.Shapes) {
                        options.CancellationToken.ThrowIfCancellationRequested();
                        WriteShape(writer, contentPage, shape, options, contentProjection, textStyles, contentVisibility, ref foreignImageIndex);
                    }
                    VisioRenderLabelLayout? labelLayout = options.ResolveConnectorLabelOverlaps
                        ? VisioRenderLabelLayout.Create(contentPage, contentVisibility) : null;
                    foreach (VisioConnector connector in contentPage.Connectors) {
                        options.CancellationToken.ThrowIfCancellationRequested();
                        if (!contentVisibility.IsVisible(connector)) continue;
                        WriteConnector(writer, contentPage, connector, options, contentProjection, labelLayout, textStyles);
                    }
                    if (background) writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }

            options.CancellationToken.ThrowIfCancellationRequested();
            return builder.ToString();
        }

        private static void WriteEmbeddedFonts(
            XmlWriter writer,
            OfficeFontFaceCollection fonts,
            System.Threading.CancellationToken cancellationToken) {
            if (fonts == null || fonts.Faces.Count == 0) return;
            var css = new StringBuilder();
            foreach (OfficeFontFace face in fonts.Faces) {
                cancellationToken.ThrowIfCancellationRequested();
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

        private static void WriteShape(XmlWriter writer, VisioPage page, VisioShape shape, VisioSvgSaveOptions options, VisioRenderProjection projection, VisioNativeTextStyleResolver textStyles, VisioRenderLayerVisibility layerVisibility, ref int foreignImageIndex) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!layerVisibility.IsVisible(shape)) {
                foreach (VisioShape child in shape.Children) {
                    WriteShape(writer, page, child, options, projection, textStyles, layerVisibility, ref foreignImageIndex);
                }
                return;
            }
            VisioNativeShapeTransform transform;
            try {
                transform = VisioNativeShapeTransform.Create(shape, options.ImageDiagnostics, options.ImageDiagnosticSource);
            } catch (Exception exception) when (exception is ArgumentException || exception is InvalidDataException) {
                VisioNativeShapeTransform.ReportInvalid(shape, options.ImageDiagnostics, options.ImageDiagnosticSource, exception);
                return;
            }
            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-visio-shape-id", shape.Id);
            if (!string.IsNullOrWhiteSpace(shape.NameU)) {
                writer.WriteAttributeString("data-visio-nameu", shape.NameU);
            }

            bool foreign = VisioForeignImage.IsForeign(shape);
            if (foreign) {
                if (VisioForeignImage.TryGetProjection(shape, page, projection.GeometryDensity, options.ImageDiagnostics, options.ImageDiagnosticSource, out OfficeImageProjection imageProjection)) {
                    OfficeRasterImage raster = VisioForeignImage.Decode(shape, options.ImageCodec, options.ImageDiagnostics, options.ImageDiagnosticSource, options.CancellationToken);
                    OfficeSvgImageRenderer.WriteImage(writer, SvgNamespace,
                        OfficeSvgImageRenderer.CreateDataUri("image/png", OfficePngWriter.Encode(raster, options.CancellationToken)),
                        imageProjection.Translate(0, projection.ContentOffsetY), preserveAspectRatio: "none",
                        writeAdditionalAttributes: imageWriter => imageWriter.WriteAttributeString("data-officeimo-foreign-image", "true"),
                        clipPathId: "visio-foreign-clip-" + (++foreignImageIndex).ToString(CultureInfo.InvariantCulture));
                }
            } else {
                WriteShapeGeometry(writer, page, shape, projection);
                if (transform.HasReflection && VisioShapeGeometry.ResolveRenderKind(shape) == "database"
                    && !VisioShapeGeometry.TryGetRenderClosedPaths(shape, out _))
                    VisioNativeShapeTransform.ReportArtwork(shape, options.ImageDiagnostics, options.ImageDiagnosticSource);
            }

            if (!foreign && options.RenderStencilArtwork) {
                bool preview = WritePackagePreviewArtwork(writer, page, shape, options, projection);
                if (!preview) {
                    WriteStencilArtwork(writer, page, shape, projection);
                }
                if (transform.HasReflection && (preview || !string.IsNullOrEmpty(VisioStencilArtwork.GetKey(shape))))
                    VisioNativeShapeTransform.ReportArtwork(shape, options.ImageDiagnostics, options.ImageDiagnosticSource);
            }

            if (options.RenderText && !string.IsNullOrEmpty(shape.Text)) {
                WriteShapeText(writer, page, shape, projection, options, textStyles, transform);
            }

            foreach (VisioShape child in shape.Children) {
                options.CancellationToken.ThrowIfCancellationRequested();
                WriteShape(writer, page, child, options, projection, textStyles, layerVisibility, ref foreignImageIndex);
            }

            writer.WriteEndElement();
        }


    }
}
