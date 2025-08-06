using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html.Converters {
    internal class WordToHtmlConverter {
        public async Task<string> ConvertAsync(WordDocument document, WordToHtmlOptions options) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordToHtmlOptions();

            var context = BrowsingContext.New(Configuration.Default);
            var htmlDoc = await context.OpenNewAsync();

            var head = htmlDoc.Head;
            var body = htmlDoc.Body;

            var charset = htmlDoc.CreateElement("meta");
            charset.SetAttribute("charset", "UTF-8");
            head.AppendChild(charset);

            var props = document.BuiltinDocumentProperties;
            var title = htmlDoc.CreateElement("title");
            title.TextContent = string.IsNullOrEmpty(props?.Title) ? "Document" : props.Title;
            head.AppendChild(title);

            void AddMeta(string name, string value) {
                if (!string.IsNullOrEmpty(value)) {
                    var meta = htmlDoc.CreateElement("meta");
                    meta.SetAttribute("name", name);
                    meta.SetAttribute("content", value);
                    head.AppendChild(meta);
                }
            }

            if (props != null) {
                AddMeta("author", props.Creator);
                AddMeta("description", props.Description);
                AddMeta("keywords", props.Keywords);
                AddMeta("subject", props.Subject);
            }

            if (!string.IsNullOrEmpty(options.FontFamily)) {
                body.SetAttribute("style", $"font-family:{options.FontFamily}");
            }

            Stack<IElement> listStack = new Stack<IElement>();
            Stack<IElement> itemStack = new Stack<IElement>();

            void CloseLists() {
                while (listStack.Count > 0) {
                    listStack.Pop();
                }
                while (itemStack.Count > 0) {
                    itemStack.Pop();
                }
            }

            string MimeFromFileName(string fileName) {
                var ext = Path.GetExtension(fileName)?.ToLowerInvariant();
                return ext switch {
                    ".jpg" => "image/jpeg",
                    ".jpeg" => "image/jpeg",
                    ".png" => "image/png",
                    ".gif" => "image/gif",
                    ".bmp" => "image/bmp",
                    ".tif" => "image/tiff",
                    ".tiff" => "image/tiff",
                    _ => "image/png"
                };
            }

            void AppendRuns(IElement parent, WordParagraph para) {
                foreach (var run in FormattingHelper.GetFormattedRuns(para)) {
                    if (run.Image != null) {
                        var img = htmlDoc.CreateElement("img") as IHtmlImageElement;
                        string src;
                        var imgObj = run.Image;
                        if (imgObj.IsExternal && imgObj.ExternalUri != null) {
                            src = imgObj.ExternalUri.ToString();
                        } else {
                            var bytes = imgObj.GetBytes();
                            var mime = MimeFromFileName(imgObj.FileName);
                            src = $"data:{mime};base64,{Convert.ToBase64String(bytes)}";
                        }
                        img!.Source = src;
                        parent.AppendChild(img);
                        continue;
                    }

                    if (string.IsNullOrEmpty(run.Text)) {
                        continue;
                    }

                    INode node = htmlDoc.CreateTextNode(run.Text);

                    if (run.Bold) {
                        var strong = htmlDoc.CreateElement("strong");
                        strong.AppendChild(node);
                        node = strong;
                    }

                    if (run.Italic) {
                        var em = htmlDoc.CreateElement("em");
                        em.AppendChild(node);
                        node = em;
                    }

                    if (run.Underline) {
                        var u = htmlDoc.CreateElement("u");
                        u.AppendChild(node);
                        node = u;
                    }

                    if (!string.IsNullOrEmpty(run.Hyperlink)) {
                        var a = htmlDoc.CreateElement("a");
                        a.SetAttribute("href", run.Hyperlink);
                        a.AppendChild(node);
                        node = a;
                    }

                    if (options.IncludeFontStyles && !string.IsNullOrEmpty(options.FontFamily)) {
                        var span = htmlDoc.CreateElement("span");
                        span.SetAttribute("style", $"font-family:{options.FontFamily}");
                        span.AppendChild(node);
                        node = span;
                    }

                    parent.AppendChild(node);
                }
            }

            void AppendParagraph(IElement parent, WordParagraph para) {
                int level = para.Style.HasValue ? HeadingStyleMapper.GetLevelForHeadingStyle(para.Style.Value) : 0;
                var element = htmlDoc.CreateElement(level > 0 ? $"h{level}" : "p");
                AppendRuns(element, para);
                parent.AppendChild(element);
            }

            string? GetListStyle(DocumentTraversal.ListInfo info) {
                var format = info.NumberFormat;
                if (format == NumberFormatValues.Decimal) {
                    return "decimal";
                }
                if (format == NumberFormatValues.LowerLetter) {
                    return "lower-alpha";
                }
                if (format == NumberFormatValues.UpperLetter) {
                    return "upper-alpha";
                }
                if (format == NumberFormatValues.LowerRoman) {
                    return "lower-roman";
                }
                if (format == NumberFormatValues.UpperRoman) {
                    return "upper-roman";
                }
                if (format == NumberFormatValues.Bullet) {
                    return info.LevelText switch {
                        "o" or "◦" => "circle",
                        "■" or "§" => "square",
                        _ => "disc",
                    };
                }
                return null;
            }

            foreach (var section in DocumentTraversal.EnumerateSections(document)) {
                foreach (var element in section.Elements) {
                    if (element is WordParagraph paragraph) {
                        var listInfo = DocumentTraversal.GetListInfo(paragraph);
                        if (listInfo != null) {
                            int level = listInfo.Value.Level;
                            while (listStack.Count > level) {
                                listStack.Pop();
                                itemStack.Pop();
                            }
                            while (listStack.Count <= level) {
                                bool ordered = listInfo.Value.Ordered;
                                var listTag = ordered ? "ol" : "ul";
                                var listEl = htmlDoc.CreateElement(listTag);
                                if (ordered && listInfo.Value.Start > 1) {
                                    listEl.SetAttribute("start", listInfo.Value.Start.ToString());
                                }
                                if (options.IncludeListStyles) {
                                    var css = GetListStyle(listInfo.Value);
                                    if (!string.IsNullOrEmpty(css)) {
                                        listEl.SetAttribute("style", $"list-style-type:{css}");
                                    }
                                }
                                if (itemStack.Count > 0) {
                                    itemStack.Peek().AppendChild(listEl);
                                } else {
                                    body.AppendChild(listEl);
                                }
                                listStack.Push(listEl);
                            }
                            while (itemStack.Count > level) {
                                itemStack.Pop();
                            }
                            var li = htmlDoc.CreateElement("li");
                            listStack.Peek().AppendChild(li);
                            itemStack.Push(li);
                            AppendRuns(li, paragraph);
                        } else {
                            CloseLists();
                            AppendParagraph(body, paragraph);
                        }
                    } else if (element is WordTable table) {
                        CloseLists();
                        var tableEl = htmlDoc.CreateElement("table");
                        foreach (var row in table.Rows) {
                            var tr = htmlDoc.CreateElement("tr");
                            foreach (var cell in row.Cells) {
                                var td = htmlDoc.CreateElement("td");
                                foreach (var p in cell.Paragraphs) {
                                    AppendParagraph(td, p);
                                }
                                tr.AppendChild(td);
                            }
                            tableEl.AppendChild(tr);
                        }
                        body.AppendChild(tableEl);
                    }
                }
            }

            CloseLists();

            return htmlDoc.DocumentElement.OuterHtml;
        }
    }
}

