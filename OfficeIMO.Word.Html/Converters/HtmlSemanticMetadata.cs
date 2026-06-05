using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace OfficeIMO.Word.Html {
    internal static class HtmlSemanticMetadata {
        private const string TimeDateTimeTagPrefix = "officeimo-html-time:";

        internal static void SetBlockquoteCite(WordParagraph paragraph, string? cite) {
            if (string.IsNullOrWhiteSpace(cite) || paragraph._paragraph == null) {
                return;
            }

            paragraph._paragraph.RemoveAnnotations<HtmlBlockquoteCiteAnnotation>();
            paragraph._paragraph.AddAnnotation(new HtmlBlockquoteCiteAnnotation(cite!));
        }

        internal static bool TryGetBlockquoteCite(WordParagraph paragraph, out string cite) {
            cite = string.Empty;
            var annotation = paragraph._paragraph?.Annotation<HtmlBlockquoteCiteAnnotation>();
            if (annotation == null || string.IsNullOrEmpty(annotation.Cite)) {
                return false;
            }

            cite = annotation.Cite;
            return true;
        }

        internal static void SetTimeDateTime(WordParagraph run, string? dateTime) {
            if (string.IsNullOrWhiteSpace(dateTime) || run._run == null) {
                return;
            }

            run._run.RemoveAnnotations<HtmlTimeAnnotation>();
            run._run.AddAnnotation(new HtmlTimeAnnotation(dateTime!));
            SetTimeDateTimeMetadataRun(run._run, dateTime!);
        }

        internal static bool TryGetTimeDateTime(WordParagraph run, out string dateTime) {
            dateTime = string.Empty;
            var annotation = run._run?.Annotation<HtmlTimeAnnotation>();
            if (annotation != null && !string.IsNullOrEmpty(annotation.DateTime)) {
                dateTime = annotation.DateTime;
                return true;
            }

            if (run._run != null) {
                var nextRun = run._run.NextSibling<Run>();
                if (TryGetTimeDateTimeMetadata(nextRun, out dateTime)) {
                    return true;
                }
            }

            return false;
        }

        internal static bool IsTimeDateTimeMetadataRun(WordParagraph run) =>
            TryGetTimeDateTimeMetadata(run._run, out _);

        private static void SetTimeDateTimeMetadataRun(Run run, string dateTime) {
            var tagValue = TimeDateTimeTagPrefix + Uri.EscapeDataString(dateTime);
            var metadataRun = run.NextSibling<Run>();
            if (!IsTimeDateTimeMetadataRun(metadataRun)) {
                metadataRun = new Run(new RunProperties(new Vanish()));
                run.InsertAfterSelf(metadataRun);
            }

            var text = metadataRun!.GetFirstChild<Text>();
            if (text == null) {
                metadataRun.AppendChild(new Text(tagValue) { Space = SpaceProcessingModeValues.Preserve });
            } else {
                text.Text = tagValue;
                text.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        private static bool TryGetTimeDateTimeMetadata(Run? run, out string dateTime) {
            dateTime = string.Empty;
            if (!IsTimeDateTimeMetadataRun(run)) {
                return false;
            }

            var tag = run!.GetFirstChild<Text>()?.Text;
            dateTime = Uri.UnescapeDataString(tag!.Substring(TimeDateTimeTagPrefix.Length));
            return true;
        }

        private static bool IsTimeDateTimeMetadataRun(Run? run) {
            var tag = run?.GetFirstChild<Text>()?.Text;
            return !string.IsNullOrEmpty(tag) &&
                tag!.StartsWith(TimeDateTimeTagPrefix, StringComparison.Ordinal) &&
                run!.RunProperties?.GetFirstChild<Vanish>() != null;
        }

        private sealed class HtmlTimeAnnotation {
            internal HtmlTimeAnnotation(string dateTime) {
                DateTime = dateTime;
            }

            internal string DateTime { get; }
        }

        private sealed class HtmlBlockquoteCiteAnnotation {
            internal HtmlBlockquoteCiteAnnotation(string cite) {
                Cite = cite;
            }

            internal string Cite { get; }
        }
    }
}
