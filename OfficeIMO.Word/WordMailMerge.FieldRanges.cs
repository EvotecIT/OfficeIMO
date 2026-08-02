using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordMailMerge {
        private static bool ReplaceComplexFieldRangeWithText(
            ComplexFieldFrame field,
            string value,
            Run? sourceRun) {
            IReadOnlyList<Run> fieldRuns = field.Runs;
            FieldChar beginMarker = field.BeginMarker;
            FieldChar? endMarker = field.EndMarker;
            Run? beginRun = beginMarker.Parent as Run;
            Run? endRun = endMarker?.Parent as Run;
            if (beginRun == null || endRun == null || endMarker == null) {
                return false;
            }
            int beginRunIndex = field.Runs.IndexOf(beginRun);
            int endRunIndex = field.Runs.IndexOf(endRun);
            int beginChildIndex = beginRun.ChildElements.ToList().IndexOf(beginMarker);
            int endChildIndex = endRun.ChildElements.ToList().IndexOf(endMarker);

            var replacementText = new Text(value) { Space = SpaceProcessingModeValues.Preserve };
            if (ReferenceEquals(beginRun, endRun)) {
                if (beginChildIndex < 0 || endChildIndex < beginChildIndex) return false;
                OpenXmlElement? suffix = endMarker.NextSibling();
                for (int childIndex = endChildIndex; childIndex >= beginChildIndex; childIndex--) {
                    beginRun.ChildElements[childIndex].Remove();
                }
                if (suffix != null) beginRun.InsertBefore(replacementText, suffix);
                else beginRun.Append(replacementText);
                return true;
            }

            Run replacementRun = CreateReplacementRun(value, sourceRun);
            if (beginChildIndex < 0) return false;
            for (int childIndex = beginRun.ChildElements.Count - 1; childIndex >= beginChildIndex; childIndex--) {
                beginRun.ChildElements[childIndex].Remove();
            }
            bool preserveBeginRun = beginRun.ChildElements.Any(child => child is not RunProperties);
            if (preserveBeginRun) beginRun.InsertAfterSelf(replacementRun);
            else {
                beginRun.InsertBeforeSelf(replacementRun);
                beginRun.Remove();
            }

            for (int runIndex = beginRunIndex + 1; runIndex < endRunIndex; runIndex++) {
                fieldRuns[runIndex].Remove();
            }

            if (endChildIndex < 0) return false;
            for (int childIndex = endChildIndex; childIndex >= 0; childIndex--) {
                if (endRun.ChildElements[childIndex] is RunProperties) continue;
                endRun.ChildElements[childIndex].Remove();
            }
            if (!endRun.ChildElements.Any(child => child is not RunProperties)) {
                endRun.Remove();
            }
            return true;
        }
    }
}
