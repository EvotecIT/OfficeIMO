using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;
using OfficeIMO.Word;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        static IContainer RenderParagraph(IContainer container, WordParagraph paragraph, (int Level, string Marker)? marker, PdfSaveOptions? options, Dictionary<WordParagraph, int> footnoteMap) {
            if (paragraph == null) {
                return container;
            }

            if (paragraph.Bookmark != null) {
                container = container.Section(paragraph.Bookmark.Name);
            }

            if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                var link = paragraph.Hyperlink;
                if (!string.IsNullOrEmpty(link.Anchor)) {
                    container = container.SectionLink(link.Anchor);
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
                            row.ConstantItem(indentSize).Text(marker.Value.Marker);
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
                    span = span.FontFamily(paragraph.FontFamily);
                } else if (!string.IsNullOrEmpty(options?.FontFamily)) {
                    span = span.FontFamily(options.FontFamily);
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
                container = container.SectionLink(link.Anchor);
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
                container.Canvas((canvasObj, size) => {
                    var canvasType = canvasObj.GetType();
                    var positionType = canvasType.Assembly.GetType("QuestPDF.Infrastructure.Position");
                    var ctor = positionType!.GetConstructor(new[] { typeof(float), typeof(float) });

                    (float x1, float y1) = ParsePoint(line.From?.Value ?? "0pt,0pt");
                    (float x2, float y2) = ParsePoint(line.To?.Value ?? "0pt,0pt");
                    object start = ctor!.Invoke(new object[] { x1, y1 });
                    object end = ctor.Invoke(new object[] { x2, y2 });

                    var paint = new SKPaint {
                        Style = SKPaintStyle.Stroke,
                        Color = SKColor.Parse("#" + (stroke ?? "000000")),
                        StrokeWidth = strokeWidth
                    };

                    canvasType.GetMethod("DrawLine")!.Invoke(canvasObj, new object[] { start, end, paint });
                });

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