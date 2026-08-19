using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static Action PrepareTrackedChangeProjection(WordDocument document, WordToHtmlOptions options, bool hasRevisions) {
            if (!hasRevisions) return static () => { };

            var restore = new List<Action>();
            try {
                Body? body = document._document.Body;
                if (body != null) {
                    Body projectedBody = (Body)body.CloneNode(true);
                    ApplyTrackedChangePolicy(projectedBody, options.TrackedChangePolicy);
                    document._document.Body = projectedBody;
                    restore.Add(() => document._document.Body = body);
                }

                if (options.ExportHeadersAndFooters) {
                    var projectedRegions = new List<WordHeaderFooter>();
                    foreach (WordSection section in document.Sections) {
                        ProjectHeaderFooter(section.Header.Default, projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.Header.First, projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.Header.Even, projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.Footer.Default, projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.Footer.First, projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.Footer.Even, projectedRegions, options.TrackedChangePolicy, restore);
                    }
                }

                var mainPart = document._wordprocessingDocument.MainDocumentPart;
                if (options.ExportFootnotes && mainPart?.FootnotesPart?.Footnotes is Footnotes footnotes) {
                    Footnotes projected = (Footnotes)footnotes.CloneNode(true);
                    ApplyTrackedChangePolicy(projected, options.TrackedChangePolicy);
                    mainPart.FootnotesPart.Footnotes = projected;
                    restore.Add(() => mainPart.FootnotesPart.Footnotes = footnotes);
                }
                if (options.ExportEndnotes && mainPart?.EndnotesPart?.Endnotes is Endnotes endnotes) {
                    Endnotes projected = (Endnotes)endnotes.CloneNode(true);
                    ApplyTrackedChangePolicy(projected, options.TrackedChangePolicy);
                    mainPart.EndnotesPart.Endnotes = projected;
                    restore.Add(() => mainPart.EndnotesPart.Endnotes = endnotes);
                }
                if (options.ExportComments && mainPart?.WordprocessingCommentsPart?.Comments is Comments comments) {
                    Comments projected = (Comments)comments.CloneNode(true);
                    ApplyTrackedChangePolicy(projected, options.TrackedChangePolicy);
                    mainPart.WordprocessingCommentsPart.Comments = projected;
                    restore.Add(() => mainPart.WordprocessingCommentsPart.Comments = comments);
                }
            } catch {
                RestoreProjectedRoots(restore);
                throw;
            }

            return () => RestoreProjectedRoots(restore);
        }

        private static void RestoreProjectedRoots(IReadOnlyList<Action> restore) {
            for (int index = restore.Count - 1; index >= 0; index--) restore[index]();
        }

        private static void ProjectHeaderFooter(WordHeaderFooter? region, List<WordHeaderFooter> projectedRegions,
            WordTrackedChangeExportPolicy policy, List<Action> restore) {
            if (region == null || projectedRegions.Any(existing => ReferenceEquals(existing, region))) return;
            projectedRegions.Add(region);
            if (region._header is Header header) {
                Header projected = (Header)header.CloneNode(true);
                ApplyTrackedChangePolicy(projected, policy);
                region._header = projected;
                restore.Add(() => region._header = header);
            } else if (region._footer is Footer footer) {
                Footer projected = (Footer)footer.CloneNode(true);
                ApplyTrackedChangePolicy(projected, policy);
                region._footer = projected;
                restore.Add(() => region._footer = footer);
            }
        }

        private static void ApplyTrackedChangePolicy(OpenXmlCompositeElement root, WordTrackedChangeExportPolicy policy) {
            OpenXmlElement[] revisions = root.Descendants()
                .Where(element => element.LocalName is "ins" or "del" or "moveFrom" or "moveTo")
                .Reverse()
                .ToArray();
            foreach (OpenXmlElement revision in revisions) {
                bool insertedView = revision.LocalName is "ins" or "moveTo";
                bool include = policy switch {
                    WordTrackedChangeExportPolicy.Final => insertedView,
                    WordTrackedChangeExportPolicy.Original => !insertedView,
                    WordTrackedChangeExportPolicy.Markup => true,
                    _ => throw new ArgumentOutOfRangeException(nameof(policy), policy, "Word tracked-change policy is not supported.")
                };
                if (!include) {
                    revision.Remove();
                    continue;
                }

                string? styleId = policy == WordTrackedChangeExportPolicy.Markup
                    ? insertedView ? HtmlSemanticStyleIds.InsertedText : HtmlSemanticStyleIds.DeletedText
                    : null;
                ReplaceRevisionContainer(revision, styleId);
            }

            foreach (OpenXmlElement formattingChange in root.Descendants()
                         .Where(element => element.LocalName.EndsWith("PrChange", StringComparison.Ordinal))
                         .ToArray()) {
                formattingChange.Remove();
            }
        }

        private static void ReplaceRevisionContainer(OpenXmlElement revision, string? styleId) {
            OpenXmlElement? parent = revision.Parent;
            if (parent == null) return;
            foreach (OpenXmlElement source in revision.ChildElements.ToArray()) {
                OpenXmlElement child = source.CloneNode(true);
                RestoreDeletedText(child);
                if (styleId != null) {
                    foreach (Run run in EnumerateRuns(child)) ApplyReviewStyle(run, styleId);
                }
                parent.InsertBefore(child, revision);
            }
            revision.Remove();
        }

        private static IEnumerable<Run> EnumerateRuns(OpenXmlElement element) {
            if (element is Run run) yield return run;
            foreach (Run descendant in element.Descendants<Run>()) yield return descendant;
        }

        private static void RestoreDeletedText(OpenXmlElement element) {
            foreach (DeletedText deleted in element.Descendants<DeletedText>().ToArray()) {
                var text = new Text(deleted.Text) { Space = deleted.Space };
                deleted.InsertAfterSelf(text);
                deleted.Remove();
            }
            if (element is DeletedText rootDeleted) {
                // DeletedText is normally nested in a run; this guard keeps malformed producer input deterministic.
                rootDeleted.Text = rootDeleted.Text ?? string.Empty;
            }
        }

        private static void ApplyReviewStyle(Run run, string styleId) {
            RunProperties properties = run.RunProperties ?? new RunProperties();
            RunStyle? style = properties.GetFirstChild<RunStyle>();
            if (style == null) {
                style = new RunStyle();
                properties.PrependChild(style);
            }
            style.Val = styleId;
            if (run.RunProperties == null) run.PrependChild(properties);
        }

        private static bool IsSelectedFieldLocation(WordFieldLocationKind location, WordToHtmlOptions options) => location switch {
            WordFieldLocationKind.Header or WordFieldLocationKind.Footer => options.ExportHeadersAndFooters,
            WordFieldLocationKind.Footnote => options.ExportFootnotes,
            WordFieldLocationKind.Endnote => options.ExportEndnotes,
            _ => true
        };

        private static bool IsSelectedReviewLocation(WordReviewLocationKind location, WordToHtmlOptions options) => location switch {
            WordReviewLocationKind.Header or WordReviewLocationKind.Footer => options.ExportHeadersAndFooters,
            WordReviewLocationKind.Footnote => options.ExportFootnotes,
            WordReviewLocationKind.Endnote => options.ExportEndnotes,
            _ => true
        };

        private static void AppendReviewInventories(IHtmlDocument htmlDoc, IElement body,
            IReadOnlyList<WordRevisionInfo>? reviewInfo, IReadOnlyList<WordFieldInfo>? fields, WordToHtmlOptions options) {
            if (reviewInfo?.Count > 0) AppendRevisionInventory(htmlDoc, body, reviewInfo);
            if (fields?.Count > 0) AppendFieldInventory(htmlDoc, body, fields);
        }

        private static void AppendRevisionInventory(IHtmlDocument document, IElement body,
            IReadOnlyList<WordRevisionInfo> revisions) {
            IElement section = CreateOutputElement(document, "section");
            SetOutputAttribute(document, section, "class", "officeimo-feature officeimo-revisions", "RevisionInventory:class");
            SetOutputAttribute(document, section, "data-officeimo-review-policy", "markup", "RevisionInventory:policy");
            IElement heading = CreateOutputElement(document, "h2");
            SetOutputText(document, heading, "Tracked changes", "RevisionInventory:heading");
            section.AppendChild(heading);
            IElement note = CreateOutputElement(document, "p");
            SetOutputAttribute(document, note, "class", "officeimo-diagnostic", "RevisionInventory:diagnostic-class");
            SetOutputAttribute(document, note, "data-officeimo-loss", "review-only", "RevisionInventory:loss");
            SetOutputText(document, note, "Revision content is static review markup; accept, reject, and authoring behavior remain in Word.", "RevisionInventory:diagnostic");
            section.AppendChild(note);
            IElement list = CreateOutputElement(document, "ul");
            foreach (WordRevisionInfo revision in revisions) {
                IElement item = CreateOutputElement(document, "li");
                SetOutputAttribute(document, item, "data-officeimo-revision-type", revision.RevisionType.ToString(), "RevisionInventory:type");
                if (!string.IsNullOrWhiteSpace(revision.Author)) {
                    SetOutputAttribute(document, item, "data-officeimo-author", revision.Author!, "RevisionInventory:author");
                }
                SetOutputText(document, item, revision.RevisionType + ": " + revision.AffectedText, "RevisionInventory:text");
                list.AppendChild(item);
            }
            section.AppendChild(list);
            body.AppendChild(section);
        }

        private static void AppendFieldInventory(IHtmlDocument document, IElement body,
            IReadOnlyList<WordFieldInfo> fields) {
            IElement section = CreateOutputElement(document, "section");
            SetOutputAttribute(document, section, "class", "officeimo-feature officeimo-fields", "FieldInventory:class");
            IElement heading = CreateOutputElement(document, "h2");
            SetOutputText(document, heading, "Fields", "FieldInventory:heading");
            section.AppendChild(heading);
            IElement note = CreateOutputElement(document, "p");
            SetOutputAttribute(document, note, "class", "officeimo-diagnostic", "FieldInventory:diagnostic-class");
            SetOutputAttribute(document, note, "data-officeimo-loss", "review-only", "FieldInventory:loss");
            SetOutputText(document, note, "Field instructions are inert metadata; HTML shows stored results and never evaluates Word fields.", "FieldInventory:diagnostic");
            section.AppendChild(note);
            IElement list = CreateOutputElement(document, "dl");
            foreach (WordFieldInfo field in fields) {
                IElement term = CreateOutputElement(document, "dt");
                SetOutputAttribute(document, term, "data-officeimo-field-type", field.FieldType?.ToString() ?? "Unknown", "FieldInventory:type");
                SetOutputText(document, term, field.FieldType?.ToString() ?? "Field", "FieldInventory:label");
                list.AppendChild(term);
                IElement value = CreateOutputElement(document, "dd");
                SetOutputAttribute(document, value, "data-officeimo-field-instruction", field.InstructionText, "FieldInventory:instruction");
                SetOutputAttribute(document, value, "data-officeimo-field-location", field.LocationKind.ToString(), "FieldInventory:location");
                SetOutputText(document, value, field.ResultText, "FieldInventory:result");
                list.AppendChild(value);
            }
            section.AppendChild(list);
            body.AppendChild(section);
        }
    }
}
