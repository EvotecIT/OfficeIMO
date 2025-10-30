using DocumentFormat.OpenXml.Vml;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using SkiaSharp;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Globalization;
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
                if (skStream != null) {
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
                    return;
                }

                // Fallback: try to locate system font files cross-platform
                string? path = TryResolveSystemFontFile(fontFamily!);
                if (!string.IsNullOrEmpty(path) && File.Exists(path)) {
                    using var fs = File.OpenRead(path);
                    FontManager.RegisterFontWithCustomName(fontFamily!, fs);
                }
            } catch {
            }
        }

        static string? TryResolveSystemFontFile(string family) {
            try {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                    var fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
                    if (!string.IsNullOrEmpty(fontsDir) && Directory.Exists(fontsDir)) {
                        // Common Arial files
                        string[] candidates = new[] { "arial.ttf", "arial.ttf", "ARIAL.TTF", "Arial.ttf", "arialmt.ttf", "ARIALMT.TTF" };
                        foreach (var c in candidates) {
                            var p = System.IO.Path.Combine(fontsDir, c);
                            if (File.Exists(p)) return p;
                        }
                        // Last resort: search for files containing family name
                        var file = Directory.EnumerateFiles(fontsDir, "*.ttf").Concat(Directory.EnumerateFiles(fontsDir, "*.otf")).FirstOrDefault(f => System.IO.Path.GetFileNameWithoutExtension(f).Contains(family, StringComparison.OrdinalIgnoreCase));
                        if (file != null) return file;
                    }
                } else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                    string[] paths = new[] {
                        "/System/Library/Fonts/Supplemental/Arial.ttf",
                        "/Library/Fonts/Arial.ttf",
                        "/System/Library/Fonts/Arial.ttf"
                    };
                    foreach (var p in paths) if (File.Exists(p)) return p;
                } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                    // DejaVu Sans is most common
                    string[] roots = new[] { "/usr/share/fonts", "/usr/local/share/fonts", System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts") };
                    foreach (var r in roots) {
                        if (!Directory.Exists(r)) continue;
                        var file = Directory.EnumerateFiles(r, "*", SearchOption.AllDirectories)
                            .FirstOrDefault(f => System.IO.Path.GetFileName(f).EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) && System.IO.Path.GetFileNameWithoutExtension(f).Contains(family.Replace(" ", string.Empty), StringComparison.OrdinalIgnoreCase));
                        if (file != null) return file;
                        // Specific fallback for DejaVu Sans
                        var dv = Directory.EnumerateFiles(r, "DejaVuSans.ttf", SearchOption.AllDirectories).FirstOrDefault();
                        if (dv != null && family.Contains("DejaVu", StringComparison.OrdinalIgnoreCase)) return dv;
                    }
                }
            } catch { }
            return null;
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

            // Paragraph background and borders
            if (!string.IsNullOrEmpty(paragraph.ShadingFillColorHex)) {
                container = container.Background("#" + paragraph.ShadingFillColorHex);
            }
            container = ApplyParagraphBorders(container, paragraph);

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

                var runObjs = paragraph.GetRuns().ToList();
                // Prefer run-accurate rendering when available; otherwise fall back to paragraph text
                string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
                bool hasRenderableRuns = runObjs.Count > 0 && runObjs.Any(r => r.IsImage || !string.IsNullOrEmpty(r.Text));
                if (hasRenderableRuns || !string.IsNullOrEmpty(content) || marker != null) {
                    if (marker != null) {
                        const float indentSize = 15f;
                        col.Item().Row(row => {
                            if (marker.Value.Level > 0) {
                                row.ConstantItem(indentSize * marker.Value.Level);
                            }
                            row.AutoItem().Text(marker.Value.Marker + " ");
                            row.RelativeItem().Text(text => {
                                if (hasRenderableRuns) {
                                    foreach (var run in runObjs) {
                                        if (run.IsImage) continue; // images handled separately above
                                        if (string.IsNullOrEmpty(run.Text)) continue;
                                        var span = text.Span(run.Text!);
                                        // apply paragraph defaults first, then run overrides
                                        ApplyFormatting(span);
                                        ApplyRunFormatting(ref span, run, options);
                                    }
                                } else {
                                    ApplyFormatting(text.Span(content));
                                }
                                if (currentFootnoteNumber != null) {
                                    text.Span(currentFootnoteNumber.Value.ToString()).FontSize(8).Superscript();
                                }
                            });
                        });
                    } else {
                        col.Item().Text(text => {
                            if (hasRenderableRuns) {
                                foreach (var run in runObjs) {
                                    if (run.IsImage) continue; // images handled above
                                    if (string.IsNullOrEmpty(run.Text)) continue;
                                    var span = text.Span(run.Text!);
                                    // apply paragraph defaults first, then run overrides
                                    ApplyFormatting(span);
                                    ApplyRunFormatting(ref span, run, options);
                                }
                            } else {
                                ApplyFormatting(text.Span(content));
                            }
                            if (currentFootnoteNumber != null) {
                                text.Span(currentFootnoteNumber.Value.ToString()).FontSize(8).Superscript();
                            }
                        });
                    }
                }
            });

            return container;

            string? ResolveRegisteredFamily(string? name) {
                try {
                    if (!string.IsNullOrWhiteSpace(name)) {
                        if (options?.FontFilePaths != null && options.FontFilePaths.TryGetValue(name!, out var path) && File.Exists(path)) {
                            using var tf = SKTypeface.FromFile(path);
                            return tf?.FamilyName ?? name;
                        }
                        if (options?.FontStreams != null && options.FontStreams.TryGetValue(name!, out var s) && s != null) {
                            Stream src = s;
                            if (src.CanSeek) src.Position = 0;
                            using MemoryStream ms = new();
                            src.CopyTo(ms);
                            ms.Position = 0;
                            using var tf = SKTypeface.FromStream(new SKManagedStream(ms));
                            if (src.CanSeek) src.Position = 0;
                            return tf?.FamilyName ?? name;
                        }
                    }
                } catch { }
                return name;
            }

            void ApplyFormatting(TextSpanDescriptor span) {
                if (!string.IsNullOrEmpty(paragraph.FontFamily)) {
                    var fam = ResolveRegisteredFamily(paragraph.FontFamily!);
                    EmbedFont(fam);
                    span = span.FontFamily(fam!);
                } else if (!string.IsNullOrEmpty(options?.FontFamily)) {
                    var defFont = ResolveRegisteredFamily(options!.FontFamily!);
                    EmbedFont(defFont);
                    span = span.FontFamily(defFont!);
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

            static string? MapHighlight(W.HighlightColorValues? highlight) {
                if (!highlight.HasValue) return null;
                var v = highlight.Value;
                if (v == W.HighlightColorValues.None) return null;
                if (v == W.HighlightColorValues.Black) return "#000000";
                if (v == W.HighlightColorValues.Blue) return "#0000ff";
                if (v == W.HighlightColorValues.Cyan) return "#00ffff";
                if (v == W.HighlightColorValues.Green) return "#00ff00";
                if (v == W.HighlightColorValues.Magenta) return "#ff00ff";
                if (v == W.HighlightColorValues.Red) return "#ff0000";
                if (v == W.HighlightColorValues.Yellow) return "#ffff00";
                if (v == W.HighlightColorValues.White) return "#ffffff";
                if (v == W.HighlightColorValues.DarkBlue) return "#00008b";
                if (v == W.HighlightColorValues.DarkCyan) return "#008b8b";
                if (v == W.HighlightColorValues.DarkGreen) return "#006400";
                if (v == W.HighlightColorValues.DarkMagenta) return "#8b008b";
                if (v == W.HighlightColorValues.DarkRed) return "#8b0000";
                if (v == W.HighlightColorValues.DarkYellow) return "#b8860b";
                if (v == W.HighlightColorValues.LightGray) return "#d3d3d3";
                if (v == W.HighlightColorValues.DarkGray) return "#a9a9a9";
                return null;
            }

            void ApplyRunFormatting(ref TextSpanDescriptor span, WordParagraph run, PdfSaveOptions? opt) {
                if (string.IsNullOrEmpty(run.Text)) return;
                if (run.Bold) span = span.Bold();
                if (run.Italic) span = span.Italic();
                if (run.Underline != null) span = span.Underline();
                if (run.Strike || run.DoubleStrike) span = span.Strikethrough();
                if (run.VerticalTextAlignment == W.VerticalPositionValues.Superscript) span = span.Superscript();
                if (run.VerticalTextAlignment == W.VerticalPositionValues.Subscript) span = span.Subscript();
                // Inline hyperlink on text spans is not supported by QuestPDF directly.
                // Paragraph-level hyperlinks are applied earlier; skip span-level link here.
                // Monospace/code detection via run font or provided default
                string? mono = null;
                if (!string.IsNullOrEmpty(run.FontFamily)) mono = run.FontFamily;
                mono ??= FontResolver.Resolve("monospace") ?? opt?.FontFamily;
                if (!string.IsNullOrEmpty(mono)) {
                    var fam = ResolveRegisteredFamily(mono)!;
                    EmbedFont(fam);
                    span = span.FontFamily(fam);
                }
                if (!string.IsNullOrEmpty(run.ColorHex)) span = span.FontColor("#" + run.ColorHex);
                var hl = MapHighlight(run.Highlight);
                if (!string.IsNullOrEmpty(hl)) span = span.BackgroundColor(hl!);
            }

            static IContainer ApplyParagraphBorders(IContainer cont, WordParagraph p) {
                var b = p.Borders;
                if (b == null) return cont;

                // Determine a uniform color if possible
                var colors = new List<string?> { b.TopColorHex, b.BottomColorHex, b.LeftColorHex, b.RightColorHex };
                colors.RemoveAll(string.IsNullOrEmpty);
                if (colors.Count > 0 && colors.Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1) {
                    cont = cont.BorderColor("#" + colors[0]!);
                }

                float BorderWidth(uint? size) => size.HasValue ? size.Value / 8f : 0f;
                if (b.TopStyle != null && b.TopStyle != W.BorderValues.Nil && b.TopStyle != W.BorderValues.None) cont = cont.BorderTop(BorderWidth(b.TopSize?.Value));
                if (b.BottomStyle != null && b.BottomStyle != W.BorderValues.Nil && b.BottomStyle != W.BorderValues.None) cont = cont.BorderBottom(BorderWidth(b.BottomSize?.Value));
                if (b.LeftStyle != null && b.LeftStyle != W.BorderValues.Nil && b.LeftStyle != W.BorderValues.None) cont = cont.BorderLeft(BorderWidth(b.LeftSize?.Value));
                if (b.RightStyle != null && b.RightStyle != W.BorderValues.Nil && b.RightStyle != W.BorderValues.None) cont = cont.BorderRight(BorderWidth(b.RightSize?.Value));
                return cont;
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
