using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;
using OfficeIMO.Word;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        static readonly HashSet<string> _embeddedFonts = new();

        static void EmbedFont(string? fontFamily) {
            if (string.IsNullOrWhiteSpace(fontFamily)) {
                return;
            }
            if (!_embeddedFonts.Add(fontFamily!)) {
                return;
            }
            try {
                using SKTypeface? typeface = SKTypeface.FromFamilyName(fontFamily);
                using SKStreamAsset? skStream = typeface?.OpenStream();
                if (skStream == null) {
                    return;
                }
                using MemoryStream ms = new();
                if (skStream.HasLength) {
                    byte[] buffer = new byte[skStream.Length];
                    skStream.Read(buffer, buffer.Length);
                    ms.Write(buffer, 0, buffer.Length);
                } else {
                    byte[] buffer = new byte[4096];
                    int read;
                    while ((read = skStream.Read(buffer, buffer.Length)) > 0) {
                        ms.Write(buffer, 0, read);
                    }
                }
                ms.Position = 0;
                FontManager.RegisterFontWithCustomName(fontFamily!, ms);
            } catch {
            }
        }
        static void RenderElement(ColumnDescriptor column, WordElement element, Func<WordParagraph, (int Level, string Marker)?> getMarker, PdfSaveOptions? options, Dictionary<WordParagraph, int> footnoteMap) {
            switch (element) {
                case WordParagraph paragraph:
                    column.Item().Element(e => RenderParagraph(e, paragraph, getMarker(paragraph), options, footnoteMap));
                    break;
                case WordTable table:
                    column.Item().Element(e => RenderTable(e, table, getMarker, options, footnoteMap));
                    break;
                case WordImage image:
                    column.Item().Element(e => RenderImage(e, image));
                    break;
                case WordHyperLink link:
                    column.Item().Element(e => RenderHyperLink(e, link));
                    break;
                case WordShape shape:
                    column.Item().Element(e => RenderShape(e, shape));
                    break;
            }
        }
        static IContainer RenderParagraph(IContainer container, WordParagraph paragraph, (int Level, string Marker)? marker, PdfSaveOptions? options, Dictionary<WordParagraph, int> footnoteMap) {
            if (paragraph == null) {
                return container;
            }

            if (!string.IsNullOrEmpty(paragraph.Bookmark?.Name)) {
                container = container.Section(paragraph.Bookmark!.Name!);
            }

            if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                var link = paragraph.Hyperlink;
                if (!string.IsNullOrEmpty(link.Anchor)) {
                    container = container.SectionLink(link.Anchor!);
                } else if (link.Uri != null) {
                    container = container.Hyperlink(link.Uri.ToString());
                }
            }

            if (paragraph.ParagraphAlignment == W.JustificationValues.Center) {
                container = container.AlignCenter();
            } else if (paragraph.ParagraphAlignment == W.JustificationValues.Right) {
                container = container.AlignRight();
            } else if (paragraph.ParagraphAlignment == W.JustificationValues.Both) {
                container = container.AlignLeft();
            }

            int? currentFootnoteNumber = null;
            if (footnoteMap.TryGetValue(paragraph, out int num)) {
                currentFootnoteNumber = num;
            }

            container.Column(col => {
                if (paragraph.Image != null) {
                    col.Item().Element(imageContainer => {
                        var img = paragraph.Image;
                        var sized = imageContainer;
                        if (img.Width.HasValue) {
                            sized = sized.Width((float)(img.Width.Value * 72 / 96));
                        }
                        if (img.Height.HasValue) {
                            sized = sized.Height((float)(img.Height.Value * 72 / 96));
                        }
                        sized.Image(ImageEmbedder.GetImageBytes(img));
                    });
                }

                string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
                if (!string.IsNullOrEmpty(content) || marker != null) {
                    if (marker != null) {
                        const float indentSize = 15f;
                        col.Item().Row(row => {
                            if (marker.Value.Level > 0) {
                                row.ConstantItem(indentSize * marker.Value.Level);
                            }
                            row.AutoItem().Text(marker.Value.Marker + " ");
                            row.RelativeItem().Text(text => {
                                ApplyFormatting(text.Span(content));
                                if (currentFootnoteNumber != null) {
                                    text.Span(currentFootnoteNumber.Value.ToString()).FontSize(8).Superscript();
                                }
                            });
                        });
                    } else {
                        col.Item().Text(text => {
                            ApplyFormatting(text.Span(content));
                            if (currentFootnoteNumber != null) {
                                text.Span(currentFootnoteNumber.Value.ToString()).FontSize(8).Superscript();
                            }
                        });
                    }
                }
            });

            return container;

            void ApplyFormatting(TextSpanDescriptor span) {
                if (!string.IsNullOrEmpty(paragraph.FontFamily)) {
                    EmbedFont(paragraph.FontFamily);
                    span = span.FontFamily(paragraph.FontFamily!);
                } else if (!string.IsNullOrEmpty(options?.FontFamily)) {
                    var defFont = options!.FontFamily!;
                    EmbedFont(defFont);
                    span = span.FontFamily(defFont);
                }
                if (paragraph.Bold) {
                    span = span.Bold();
                }
                if (paragraph.Italic) {
                    span = span.Italic();
                }
                if (paragraph.Underline != null) {
                    span = span.Underline();
                }
                if (!string.IsNullOrEmpty(paragraph.ColorHex)) {
                    span = span.FontColor("#" + paragraph.ColorHex);
                }
                if (paragraph.Style.HasValue) {
                    switch (paragraph.Style.Value) {
                        case WordParagraphStyles.Heading1:
                            span.FontSize(24).Bold();
                            break;
                        case WordParagraphStyles.Heading2:
                            span.FontSize(20).Bold();
                            break;
                        case WordParagraphStyles.Heading3:
                            span.FontSize(16).Bold();
                            break;
                        case WordParagraphStyles.Heading4:
                            span.FontSize(14).Bold();
                            break;
                        case WordParagraphStyles.Heading5:
                            span.FontSize(13).Bold();
                            break;
                        case WordParagraphStyles.Heading6:
                            span.FontSize(12).Bold();
                            break;
                    }
                }
            }
        }

        static IContainer RenderImage(IContainer container, WordImage image) {
            if (image == null) {
                return container;
            }

            var sized = container;
            if (image.Width.HasValue) {
                sized = sized.Width((float)(image.Width.Value * 72 / 96));
            }
            if (image.Height.HasValue) {
                sized = sized.Height((float)(image.Height.Value * 72 / 96));
            }
            sized.Image(ImageEmbedder.GetImageBytes(image));

            return container;
        }

        static IContainer RenderHyperLink(IContainer container, WordHyperLink link) {
            if (link == null) {
                return container;
            }

            if (!string.IsNullOrEmpty(link.Anchor)) {
                container = container.SectionLink(link.Anchor!);
            } else if (link.Uri != null) {
                container = container.Hyperlink(link.Uri.ToString());
            }

            container.Text(link.Text);

            return container;
        }

        static IContainer RenderShape(IContainer container, WordShape shape) {
            if (shape == null) {
                return container;
            }
            float width = (float)shape.Width;
            float height = (float)shape.Height;

            string? fill = shape.FillColorHex;
            string? stroke = shape.StrokeColorHex;
            float strokeWidth = (float)(shape.StrokeWeight ?? 1);
            bool drawStroke = shape.Stroked ?? false;

            var type = shape.GetType();
            var runField = type.GetField("_run", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var run = runField?.GetValue(shape) as DocumentFormat.OpenXml.Wordprocessing.Run;
            string text = run?.InnerText ?? string.Empty;

            var lineField = type.GetField("_line", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var line = lineField?.GetValue(shape) as Line;

            if (line != null) {
                (float x1, float y1) = ParsePoint(line.From?.Value ?? "0pt,0pt");
                (float x2, float y2) = ParsePoint(line.To?.Value ?? "0pt,0pt");

                string svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='{width.ToString(CultureInfo.InvariantCulture)}' height='{height.ToString(CultureInfo.InvariantCulture)}' viewBox='0 0 {width.ToString(CultureInfo.InvariantCulture)} {height.ToString(CultureInfo.InvariantCulture)}'><line x1='{x1.ToString(CultureInfo.InvariantCulture)}' y1='{y1.ToString(CultureInfo.InvariantCulture)}' x2='{x2.ToString(CultureInfo.InvariantCulture)}' y2='{y2.ToString(CultureInfo.InvariantCulture)}' stroke='#{stroke ?? "000000"}' stroke-width='{strokeWidth.ToString(CultureInfo.InvariantCulture)}' /></svg>";

                container.Svg(svg);

                return container;
            }

            container = container.Width(width).Height(height);

            if (!string.IsNullOrEmpty(fill)) {
                container = container.Background("#" + fill);
            }

            if (drawStroke) {
                container = container.BorderColor("#" + (stroke ?? "000000")).Border(strokeWidth);
            }

            if (!string.IsNullOrWhiteSpace(text)) {
                container = container.AlignCenter().AlignMiddle();
                container.Text(text);
            }

            return container;

            static (float, float) ParsePoint(string value) {
                var parts = value.Split(',');
                return (Parse(parts[0]), Parse(parts[1]));
            }

            static float Parse(string value) {
                return (float)double.Parse(value.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
            }
        }
    }
}
