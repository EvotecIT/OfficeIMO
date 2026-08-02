using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordMailMerge {
        private static bool ReplaceComplexFieldRangeWithText(
            IReadOnlyList<Run> fieldRuns,
            string value,
            Run? sourceRun) {
            Run? beginRun = null;
            FieldChar? beginMarker = null;
            int beginRunIndex = -1;
            int beginChildIndex = -1;
            Run? endRun = null;
            FieldChar? endMarker = null;
            int endRunIndex = -1;
            int endChildIndex = -1;
            bool insideField = false;

            for (int runIndex = 0; runIndex < fieldRuns.Count; runIndex++) {
                Run run = fieldRuns[runIndex];
                for (int childIndex = 0; childIndex < run.ChildElements.Count; childIndex++) {
                    if (!(run.ChildElements[childIndex] is FieldChar fieldChar)) continue;
                    FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                    if (!insideField && fieldCharType == FieldCharValues.Begin) {
                        beginRun = run;
                        beginMarker = fieldChar;
                        beginRunIndex = runIndex;
                        beginChildIndex = childIndex;
                        insideField = true;
                        continue;
                    }
                    if (!insideField || fieldCharType != FieldCharValues.End) continue;
                    endRun = run;
                    endMarker = fieldChar;
                    endRunIndex = runIndex;
                    endChildIndex = childIndex;
                    break;
                }
                if (endMarker != null) break;
            }

            if (beginRun == null || beginMarker == null || endRun == null || endMarker == null) {
                return false;
            }

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
