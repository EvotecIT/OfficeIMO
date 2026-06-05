namespace OfficeIMO.Word.Html {
    internal static class HtmlSemanticMetadata {
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
        }

        internal static bool TryGetTimeDateTime(WordParagraph run, out string dateTime) {
            dateTime = string.Empty;
            var annotation = run._run?.Annotation<HtmlTimeAnnotation>();
            if (annotation == null || string.IsNullOrEmpty(annotation.DateTime)) {
                return false;
            }

            dateTime = annotation.DateTime;
            return true;
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
