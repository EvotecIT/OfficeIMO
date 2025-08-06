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

            bool IsOrdered(WordListStyle? style) {
                if (style == null) return true;
                string name = style.Value.ToString();
                return name.IndexOf("Bullet", StringComparison.OrdinalIgnoreCase) < 0;
            }

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
                foreach (var run in para.GetRuns()) {
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

                    if (run.Underline != null) {
                        var u = htmlDoc.CreateElement("u");
                        u.AppendChild(node);
                        node = u;
                    }

                    if (run.IsHyperLink && run.Hyperlink != null) {
                        var a = htmlDoc.CreateElement("a");
                        a.SetAttribute("href", run.Hyperlink.Uri.ToString());
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
                var element = htmlDoc.CreateElement(
                    para.Style >= WordParagraphStyles.Heading1 && para.Style <= WordParagraphStyles.Heading9
                        ? $"h{para.Style.Value - WordParagraphStyles.Heading1 + 1}"
                        : "p");
                AppendRuns(element, para);
                parent.AppendChild(element);
            }

            foreach (var section in document.Sections) {
                foreach (var element in section.Elements) {
                    if (element is WordParagraph paragraph) {
                        if (paragraph.IsListItem) {
                            int level = paragraph.ListItemLevel ?? 0;
                            while (listStack.Count > level) {
                                listStack.Pop();
                                itemStack.Pop();
                            }
                            while (listStack.Count <= level) {
                                bool ordered = IsOrdered(paragraph.ListStyle);
                                var listTag = ordered ? "ol" : "ul";
                                var listEl = htmlDoc.CreateElement(listTag);
                                if (options.IncludeListStyles) {
                                    listEl.SetAttribute("style", ordered ? "list-style-type:decimal" : "list-style-type:disc");
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

