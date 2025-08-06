using OfficeIMO.Word;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        static IContainer ApplyCellStyle(IContainer container, WordTableCell cell) {
            if (!string.IsNullOrEmpty(cell.ShadingFillColorHex)) {
                container = container.Background("#" + cell.ShadingFillColorHex);
            }

            WordTableCellBorder borders = cell.Borders;

            List<string> colors = new() {
                borders.TopColorHex,
                borders.BottomColorHex,
                borders.LeftColorHex,
                borders.RightColorHex
            };
            colors.RemoveAll(string.IsNullOrEmpty);
            if (colors.Count > 0 && colors.Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1) {
                container = container.BorderColor("#" + colors[0]);
            }

            if (HasBorder(borders.TopStyle)) {
                container = container.BorderTop(GetBorderWidth(borders.TopSize));
            }
            if (HasBorder(borders.BottomStyle)) {
                container = container.BorderBottom(GetBorderWidth(borders.BottomSize));
            }
            if (HasBorder(borders.LeftStyle)) {
                container = container.BorderLeft(GetBorderWidth(borders.LeftSize));
            }
            if (HasBorder(borders.RightStyle)) {
                container = container.BorderRight(GetBorderWidth(borders.RightSize));
            }

            return container;
        }

        static bool HasBorder(W.BorderValues? style) => style != null && style != W.BorderValues.Nil && style != W.BorderValues.None;

        static float GetBorderWidth(UInt32Value size) => size != null ? size.Value / 8f : 1f;

        static IContainer RenderParagraph(IContainer container, WordParagraph paragraph, (int Level, string Marker)? marker) {
            if (paragraph == null) {
                return container;
            }

            if (paragraph.ParagraphAlignment == W.JustificationValues.Center) {
                container = container.AlignCenter();
            } else if (paragraph.ParagraphAlignment == W.JustificationValues.Right) {
                container = container.AlignRight();
            } else if (paragraph.ParagraphAlignment == W.JustificationValues.Both) {
                container = container.AlignLeft();
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
                                if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                                    ApplyFormatting(text.Hyperlink(content, paragraph.Hyperlink.Uri.ToString()));
                                } else {
                                    ApplyFormatting(text.Span(content));
                                }
                            });
                        });
                    } else {
                        col.Item().Text(text => {
                            if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                                ApplyFormatting(text.Hyperlink(content, paragraph.Hyperlink.Uri.ToString()));
                            } else {
                                ApplyFormatting(text.Span(content));
                            }
                        });
                    }
                }
            });

            return container;

            void ApplyFormatting(TextSpanDescriptor span) {
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

            container.Text(text => {
                text.Hyperlink(link.Text, link.Uri.ToString());
            });

            return container;
        }

        static IContainer RenderShape(IContainer container, WordShape shape) {
            if (shape == null) {
                return container;
            }

            container.Text("[Shape]");
            return container;
        }
    }
}
