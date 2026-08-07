using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private bool TryReadObject(IElement token) {
            string? kindValue = GetAttribute(token, "data-officeimo-rtf-object");
            if (string.IsNullOrWhiteSpace(kindValue)) {
                return false;
            }

            RtfObject rtfObject = CreateObject(token, kindValue!);
            AddObjectResult(rtfObject, GetAttribute(token, "data-officeimo-rtf-object-result"));
            rtfObject.ResultImage = ReadImageMetadata(token, "data-officeimo-rtf-object-result-image");
            AddObject(rtfObject, token);
            return true;
        }

        private bool TryReadShape(IElement token) {
            if (!IsTrue(GetAttribute(token, "data-officeimo-rtf-shape"))) {
                return false;
            }

            var shape = new RtfShape();
            AddShapeInstructions(shape, GetAttribute(token, "data-officeimo-rtf-shape-instructions"));
            AddShapeProperties(shape, GetAttribute(token, "data-officeimo-rtf-shape-properties"));
            AddShapeText(shape, GetAttribute(token, "data-officeimo-rtf-shape-text"));
            AddShape(shape, token);
            return true;
        }

        private RtfObject CreateObject(IElement token, string kindValue) {
            var rtfObject = new RtfObject(ParseObjectKind(kindValue), DecodeBytes(GetAttribute(token, "data-officeimo-rtf-object-data"))) {
                ClassName = GetAttribute(token, "data-officeimo-rtf-object-class"),
                Name = GetAttribute(token, "data-officeimo-rtf-object-name"),
                Width = ReadIntegerAttribute(token, "data-officeimo-rtf-object-width"),
                Height = ReadIntegerAttribute(token, "data-officeimo-rtf-object-height"),
                ScaleX = ReadIntegerAttribute(token, "data-officeimo-rtf-object-scale-x"),
                ScaleY = ReadIntegerAttribute(token, "data-officeimo-rtf-object-scale-y")
            };

            return rtfObject;
        }

        private void AddObjectResult(RtfObject rtfObject, string? encodedContent) {
            string? html = DecodeString(encodedContent);
            if (string.IsNullOrEmpty(html)) {
                return;
            }

            HtmlToRtfOptions nestedOptions = CreateNestedOptions();
            HtmlToRtfResult nestedResult = HtmlConversionDocument.Parse(html!).ToRtfDocumentResult(nestedOptions);
            RtfDocument resultDocument = nestedResult.RequireValue();
            PropagateNestedDiagnostics(nestedResult.RtfDiagnostics);
            RtfParagraph? paragraph = resultDocument.Paragraphs.FirstOrDefault();
            if (paragraph != null) {
                CopyParagraphInlines(paragraph, rtfObject.Result, resultDocument);
            }
        }

        private void AddShapeText(RtfShape shape, string? encodedContent) {
            string? html = DecodeString(encodedContent);
            if (string.IsNullOrEmpty(html)) {
                return;
            }

            HtmlToRtfOptions nestedOptions = CreateNestedOptions();
            HtmlToRtfResult nestedResult = HtmlConversionDocument.Parse(html!).ToRtfDocumentResult(nestedOptions);
            RtfDocument textDocument = nestedResult.RequireValue();
            PropagateNestedDiagnostics(nestedResult.RtfDiagnostics);
            foreach (RtfParagraph paragraph in textDocument.Paragraphs) {
                RtfParagraph textParagraph = shape.AddTextBoxParagraph();
                CopyParagraphInlines(paragraph, textParagraph, textDocument);
            }
        }

        private HtmlToRtfOptions CreateNestedOptions() {
            HtmlToRtfOptions nestedOptions = _options.Clone();
            return nestedOptions;
        }

        private void PropagateNestedDiagnostics(IEnumerable<HtmlRtfConversionDiagnostic> diagnostics) {
            foreach (HtmlRtfConversionDiagnostic diagnostic in diagnostics) {
                _options.AddDiagnostic(diagnostic);
            }
        }

        private static void AddShapeInstructions(RtfShape shape, string? encodedInstructions) {
            string? text = DecodeString(encodedInstructions);
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            foreach (string line in SplitMetadataLines(text!)) {
                string[] parts = line.Split('|');
                if (parts.Length < 3) {
                    continue;
                }

                string? name = DecodeString(parts[0]);
                if (string.IsNullOrWhiteSpace(name)) {
                    continue;
                }

                int? parameter = null;
                if (!string.IsNullOrWhiteSpace(parts[1]) &&
                    int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)) {
                    parameter = parsed;
                }

                shape.AddInstruction(name!, parameter, parts[2] == "1");
            }
        }

        private static void AddShapeProperties(RtfShape shape, string? encodedProperties) {
            string? text = DecodeString(encodedProperties);
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            foreach (string line in SplitMetadataLines(text!)) {
                string[] parts = line.Split('|');
                if (parts.Length < 2) {
                    continue;
                }

                string? name = DecodeString(parts[0]);
                if (string.IsNullOrWhiteSpace(name)) {
                    continue;
                }

                shape.AddProperty(name!, DecodeString(parts[1]) ?? string.Empty);
            }
        }

        private RtfImage? ReadImageMetadata(IElement token, string prefix) {
            byte[] data = DecodeBytes(GetAttribute(token, prefix + "-data"));
            if (data.Length == 0) {
                return null;
            }

            RtfImage image = new RtfImage(ParseImageFormat(GetAttribute(token, prefix + "-format")), data) {
                Description = GetAttribute(token, prefix + "-description"),
                SourceWidth = ReadIntegerAttribute(token, prefix + "-source-width"),
                SourceHeight = ReadIntegerAttribute(token, prefix + "-source-height"),
                DesiredWidthTwips = ReadIntegerAttribute(token, prefix + "-desired-width"),
                DesiredHeightTwips = ReadIntegerAttribute(token, prefix + "-desired-height")
            };

            return image;
        }

        private void AddObject(RtfObject rtfObject, IElement token) {
            if (IsBlockCarrier(token)) {
                RtfObject blockObject = _document.AddObject(rtfObject.Kind, rtfObject.Data);
                CopyObject(rtfObject, blockObject);
                AddSectionBlock(blockObject);
                return;
            }

            RtfObject inlineObject = EnsureInlineParagraph().AddObject(rtfObject.Kind, rtfObject.Data);
            CopyObject(rtfObject, inlineObject);
        }

        private void AddShape(RtfShape shape, IElement token) {
            if (IsBlockCarrier(token)) {
                RtfShape blockShape = _document.AddShape();
                CopyShape(shape, blockShape);
                AddSectionBlock(blockShape);
                return;
            }

            RtfShape inlineShape = EnsureInlineParagraph().AddShape();
            CopyShape(shape, inlineShape);
        }

        private bool IsBlockCarrier(IElement token) {
            return string.Equals(token.LocalName, "div", StringComparison.OrdinalIgnoreCase) &&
                   _cell == null &&
                   _paragraph == null;
        }

        private void CopyObject(RtfObject source, RtfObject target) {
            target.ClassName = source.ClassName;
            target.Name = source.Name;
            target.Width = source.Width;
            target.Height = source.Height;
            target.ScaleX = source.ScaleX;
            target.ScaleY = source.ScaleY;
            target.ResultImage = source.ResultImage;
            CopyParagraphInlines(source.Result, target.Result, _document);
        }

        private void CopyShape(RtfShape source, RtfShape target) {
            foreach (RtfShapeInstruction instruction in source.Instructions) {
                target.AddInstruction(instruction.Name, instruction.Parameter, instruction.HasParameter);
            }

            foreach (RtfShapeProperty property in source.Properties) {
                target.AddProperty(property.Name, property.Value);
            }

            foreach (RtfParagraph paragraph in source.TextBoxParagraphs) {
                RtfParagraph targetParagraph = target.AddTextBoxParagraph();
                CopyParagraphInlines(paragraph, targetParagraph, _document);
            }
        }

        private static RtfObjectKind ParseObjectKind(string value) {
            switch (value.Trim().ToLowerInvariant()) {
                case "embedded":
                case "objemb":
                    return RtfObjectKind.Embedded;
                case "linked":
                case "objlink":
                    return RtfObjectKind.Linked;
                case "auto-linked":
                case "autolinked":
                case "objautlink":
                    return RtfObjectKind.AutoLinked;
                case "subscription":
                case "objsub":
                    return RtfObjectKind.Subscription;
                case "publisher":
                case "objpub":
                    return RtfObjectKind.Publisher;
                case "icon-embedded":
                case "iconembedded":
                case "objicemb":
                    return RtfObjectKind.IconEmbedded;
                default:
                    return RtfObjectKind.Unknown;
            }
        }

        private static RtfImageFormat ParseImageFormat(string? value) {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant()) {
                case "png":
                    return RtfImageFormat.Png;
                case "jpeg":
                case "jpg":
                    return RtfImageFormat.Jpeg;
                case "dib":
                    return RtfImageFormat.Dib;
                case "wmf":
                    return RtfImageFormat.Wmf;
                case "emf":
                    return RtfImageFormat.Emf;
                default:
                    return RtfImageFormat.Unknown;
            }
        }

        private static byte[] DecodeBytes(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return Array.Empty<byte>();
            }

            try {
                return Convert.FromBase64String(value!);
            } catch (FormatException) {
                return Array.Empty<byte>();
            }
        }

        private static string? DecodeString(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            try {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value!));
            } catch (FormatException) {
                return null;
            }
        }

        private static IEnumerable<string> SplitMetadataLines(string value) {
            return value.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static bool IsTrue(string? value) {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "1", StringComparison.OrdinalIgnoreCase);
        }
    }
}
