using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal static class HtmlSemanticMetadata {
        private const string TimeDateTimeTagPrefix = "officeimo-html-time:";

        internal static void SetBlockquoteCite(WordParagraph paragraph, string? cite) {
            OfficeXmlNode node = paragraph.Node();
            if (string.IsNullOrWhiteSpace(cite) || node.IsNone) {
                return;
            }

            node.RemoveAnnotations<HtmlBlockquoteCiteAnnotation>();
            node.AddAnnotation(new HtmlBlockquoteCiteAnnotation(cite!));
        }

        internal static void MarkBlockquote(WordParagraph paragraph) {
            OfficeXmlNode node = paragraph.Node();
            if (node.IsNone || node.Annotation<HtmlBlockquoteAnnotation>() != null) {
                return;
            }

            node.AddAnnotation(new HtmlBlockquoteAnnotation());
        }

        /// <summary>Records that a class-styled paragraph came from a <c>dt</c> or <c>dd</c> element.</summary>
        internal static void MarkDefinitionListItem(WordParagraph paragraph, bool description) {
            OfficeXmlNode node = paragraph.Node();
            if (node.IsNone) return;
            node.RemoveAnnotations<HtmlDefinitionListItemAnnotation>();
            node.AddAnnotation(new HtmlDefinitionListItemAnnotation(description));
        }

        /// <summary>The definition-list role a class-styled paragraph was imported with, or null.</summary>
        internal static bool? GetDefinitionListItemRole(WordParagraph paragraph) =>
            paragraph.Node().Annotation<HtmlDefinitionListItemAnnotation>()?.IsDescription;

        internal static bool IsMarkedBlockquote(WordParagraph paragraph) =>
            paragraph.Node().Annotation<HtmlBlockquoteAnnotation>() != null;

        internal static bool TryGetBlockquoteCite(WordParagraph paragraph, out string cite) {
            cite = string.Empty;
            var annotation = paragraph.Node().Annotation<HtmlBlockquoteCiteAnnotation>();
            if (annotation == null || string.IsNullOrEmpty(annotation.Cite)) {
                return false;
            }

            cite = annotation.Cite;
            return true;
        }

        internal static void SetTimeDateTime(WordParagraph run, string? dateTime) {
            OfficeXmlNode runNode = run.RunNode();
            if (string.IsNullOrWhiteSpace(dateTime) || runNode.IsNone) {
                return;
            }

            runNode.RemoveAnnotations<HtmlTimeAnnotation>();
            runNode.AddAnnotation(new HtmlTimeAnnotation(dateTime!));
            SetTimeDateTimeMetadataRun(runNode, dateTime!);
        }

        internal static bool TryGetTimeDateTime(WordParagraph run, out string dateTime) {
            dateTime = string.Empty;
            OfficeXmlNode runNode = run.RunNode();
            var annotation = runNode.Annotation<HtmlTimeAnnotation>();
            if (annotation != null && !string.IsNullOrEmpty(annotation.DateTime)) {
                dateTime = annotation.DateTime;
                return true;
            }

            if (runNode.Exists) {
                if (TryGetTimeDateTimeMetadata(runNode.NextSiblingNamed(WordXml.W, "r"), out dateTime)) {
                    return true;
                }
            }

            return false;
        }

        internal static bool IsTimeDateTimeMetadataRun(WordParagraph run) =>
            TryGetTimeDateTimeMetadata(run.RunNode(), out _);

        private static void SetTimeDateTimeMetadataRun(OfficeXmlNode run, string dateTime) {
            var tagValue = TimeDateTimeTagPrefix + Uri.EscapeDataString(dateTime);
            OfficeXmlNode metadataRun = run.NextSiblingNamed(WordXml.W, "r");
            if (!IsTimeDateTimeMetadataRun(metadataRun)) {
                OfficeXmlNode created = WordConverterNodes.CreateDetachedRun();
                OfficeXmlNode properties = created.Append(created.CreateElement(WordXml.W, "rPr"));
                properties.Append(properties.CreateElement(WordXml.W, "vanish"));
                metadataRun = run.InsertAfterSelf(created);
            }

            OfficeXmlNode text = metadataRun.WChild("t");
            if (text.IsNone) {
                metadataRun.Append(WordConverterNodes.CreateDetachedPreservedText(tagValue));
            } else {
                text.SetInnerText(tagValue);
                text.SetAttribute(WordXml.Xml, "space", "preserve");
            }
        }

        private static bool TryGetTimeDateTimeMetadata(OfficeXmlNode run, out string dateTime) {
            dateTime = string.Empty;
            if (!IsTimeDateTimeMetadataRun(run)) {
                return false;
            }

            string tag = run.WChild("t").InnerText;
            dateTime = Uri.UnescapeDataString(tag.Substring(TimeDateTimeTagPrefix.Length));
            return true;
        }

        private static bool IsTimeDateTimeMetadataRun(OfficeXmlNode run) {
            OfficeXmlNode text = run.WChild("t");
            string? tag = text.IsNone ? null : text.InnerText;
            return !string.IsNullOrEmpty(tag) &&
                tag!.StartsWith(TimeDateTimeTagPrefix, StringComparison.Ordinal) &&
                run.WChild("rPr").WChild("vanish").Exists;
        }

        private sealed class HtmlTimeAnnotation {
            internal HtmlTimeAnnotation(string dateTime) {
                DateTime = dateTime;
            }

            internal string DateTime { get; }
        }

        private sealed class HtmlBlockquoteAnnotation {
        }

        private sealed class HtmlDefinitionListItemAnnotation {
            internal HtmlDefinitionListItemAnnotation(bool isDescription) => IsDescription = isDescription;
            internal bool IsDescription { get; }
        }

        private sealed class HtmlBlockquoteCiteAnnotation {
            internal HtmlBlockquoteCiteAnnotation(string cite) {
                Cite = cite;
            }

            internal string Cite { get; }
        }
    }
}
