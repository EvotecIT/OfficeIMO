using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static Action PrepareTrackedChangeProjection(WordDocument document, WordHtmlExportOptions options, bool hasRevisions) {
            if (!hasRevisions) return static () => { };

            var restore = new List<Action>();
            try {
                OfficeXmlNode body = document.BodyNode();
                if (body.Exists) {
                    OfficeXmlNode projectedBody = body.Clone();
                    ApplyTrackedChangePolicy(projectedBody, options.TrackedChangePolicy);
                    document.SetBodyNode(projectedBody);
                    restore.Add(() => document.SetBodyNode(body));
                }

                if (options.ExportHeadersAndFooters) {
                    var projectedRegions = new Dictionary<OfficeXmlNode, OfficeXmlNode>();
                    foreach (WordSection section in document.Sections) {
                        ProjectHeaderFooter(section.PageSetup.GetHeader(WordHeaderFooterType.Default), projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.PageSetup.GetHeader(WordHeaderFooterType.First), projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.PageSetup.GetHeader(WordHeaderFooterType.Even), projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.PageSetup.GetFooter(WordHeaderFooterType.Default), projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.PageSetup.GetFooter(WordHeaderFooterType.First), projectedRegions, options.TrackedChangePolicy, restore);
                        ProjectHeaderFooter(section.PageSetup.GetFooter(WordHeaderFooterType.Even), projectedRegions, options.TrackedChangePolicy, restore);
                    }
                }

                if (options.ExportFootnotes) ProjectPartRoot(document.FootnotesNode(), options.TrackedChangePolicy, restore);
                if (options.ExportEndnotes) ProjectPartRoot(document.EndnotesNode(), options.TrackedChangePolicy, restore);
                if (options.ExportComments) ProjectPartRoot(document.CommentsNode(), options.TrackedChangePolicy, restore);
            } catch {
                RestoreProjectedRoots(restore);
                throw;
            }

            return () => RestoreProjectedRoots(restore);
        }

        private static void ProjectPartRoot(OfficeXmlNode root, WordTrackedChangeExportPolicy policy, List<Action> restore) {
            if (root.IsNone) return;
            OfficeXmlNode projected = root.Clone();
            ApplyTrackedChangePolicy(projected, policy);
            WordConverterNodes.ReplacePartRootNode(root, projected);
            restore.Add(() => WordConverterNodes.ReplacePartRootNode(projected, root));
        }

        private static void RestoreProjectedRoots(IReadOnlyList<Action> restore) {
            for (int index = restore.Count - 1; index >= 0; index--) restore[index]();
        }

        private static void ProjectHeaderFooter(WordHeaderFooter? region, Dictionary<OfficeXmlNode, OfficeXmlNode> projectedRegions,
            WordTrackedChangeExportPolicy policy, List<Action> restore) {
            if (region == null) return;
            OfficeXmlNode root = region.Node();
            if (root.IsNone) return;
            bool header = root.IsW("hdr");
            if (!projectedRegions.TryGetValue(root, out OfficeXmlNode existing)) {
                OfficeXmlNode projectedRoot = root.Clone();
                ApplyTrackedChangePolicy(projectedRoot, policy);
                if (!WordConverterNodes.ReplacePartRootNode(root, projectedRoot)) {
                    throw new InvalidOperationException(header
                        ? "The header projection has no owning part."
                        : "The footer projection has no owning part.");
                }
                restore.Add(() => WordConverterNodes.ReplacePartRootNode(projectedRoot, root));
                projectedRegions.Add(root, projectedRoot);
                projectedRegions.Add(projectedRoot, projectedRoot);
                existing = projectedRoot;
            }
            region.SetRootNode(existing);
            restore.Add(() => region.SetRootNode(root));
        }

        private static void ApplyTrackedChangePolicy(OfficeXmlNode root, WordTrackedChangeExportPolicy policy) {
            OfficeXmlNode[] revisions = root.DescendantElements()
                .Where(element => element.LocalName is "ins" or "del" or "moveFrom" or "moveTo")
                .Reverse()
                .ToArray();
            foreach (OfficeXmlNode revision in revisions) {
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

            foreach (OfficeXmlNode formattingChange in root.DescendantElements()
                         .Where(element => element.LocalName.EndsWith("PrChange", StringComparison.Ordinal))
                         .ToArray()) {
                formattingChange.Remove();
            }
        }

        private static void ReplaceRevisionContainer(OfficeXmlNode revision, string? styleId) {
            OfficeXmlNode parent = revision.Parent;
            if (parent.IsNone) return;
            foreach (OfficeXmlNode source in revision.Elements().ToArray()) {
                OfficeXmlNode child = source.Clone();
                RestoreDeletedText(child);
                if (styleId != null) {
                    foreach (OfficeXmlNode run in EnumerateRuns(child).ToArray()) ApplyReviewStyle(run, styleId);
                }
                parent.InsertBefore(child, revision);
            }
            revision.Remove();
        }

        private static IEnumerable<OfficeXmlNode> EnumerateRuns(OfficeXmlNode element) {
            if (element.IsW("r")) yield return element;
            foreach (OfficeXmlNode descendant in element.DescendantElements(WordXml.W, "r")) yield return descendant;
        }

        private static void RestoreDeletedText(OfficeXmlNode element) {
            foreach (OfficeXmlNode deleted in element.DescendantElements(WordXml.W, "delText").ToArray()) {
                OfficeXmlNode text = deleted.Parent.CreateElement(WordXml.W, "t");
                text.SetInnerText(deleted.InnerText);
                if (deleted.Attribute(WordXml.Xml, "space") is { } space) text.SetAttribute(WordXml.Xml, "space", space);
                deleted.InsertAfterSelf(text);
                deleted.Remove();
            }
            if (element.IsW("delText")) {
                // DeletedText is normally nested in a run; this guard keeps malformed producer input deterministic.
                element.SetInnerText(element.InnerText);
            }
        }

        private static void ApplyReviewStyle(OfficeXmlNode run, string styleId) {
            OfficeXmlNode properties = run.GetOrPrependChild(WordXml.W, "rPr");
            OfficeXmlNode style = properties.GetOrPrependChild(WordXml.W, "rStyle");
            style.SetAttribute(WordXml.W, "val", styleId);
        }

        private static bool IsSelectedFieldLocation(WordFieldLocationKind location, WordHtmlExportOptions options) => location switch {
            WordFieldLocationKind.Header or WordFieldLocationKind.Footer => options.ExportHeadersAndFooters,
            WordFieldLocationKind.Footnote => options.ExportFootnotes,
            WordFieldLocationKind.Endnote => options.ExportEndnotes,
            _ => true
        };

        private static bool IsSelectedReviewLocation(WordReviewLocationKind location, WordHtmlExportOptions options) => location switch {
            WordReviewLocationKind.Header or WordReviewLocationKind.Footer => options.ExportHeadersAndFooters,
            WordReviewLocationKind.Footnote => options.ExportFootnotes,
            WordReviewLocationKind.Endnote => options.ExportEndnotes,
            _ => true
        };

        private static void AppendReviewInventories(IHtmlDocument htmlDoc, IElement body,
            IReadOnlyList<WordRevisionInfo>? reviewInfo, IReadOnlyList<WordFieldInfo>? fields, WordHtmlExportOptions options) {
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
