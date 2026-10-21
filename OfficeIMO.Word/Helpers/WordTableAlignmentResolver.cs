using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Reads the logical start/end alignment tokens that the SDK's legacy table-alignment enum does not expose.</summary>
    internal static class WordTableAlignmentResolver {
        internal static TableRowAlignmentValues? Read(TableJustification? justification, bool rightToLeft) {
            string? token = justification?.Val?.InnerText;
            if (token == "start") return rightToLeft ? TableRowAlignmentValues.Right : TableRowAlignmentValues.Left;
            if (token == "end") return rightToLeft ? TableRowAlignmentValues.Left : TableRowAlignmentValues.Right;
            // Keep the normal SDK validation for unknown tokens; only the two
            // defined logical aliases need interpretation before enum parsing.
            return justification?.Val?.Value;
        }

        internal static bool IsRightToLeft(TableProperties? properties) =>
            ReadDirection(properties?.GetFirstChild<BiDiVisual>()) ?? false;

        internal static bool? ReadDirection(BiDiVisual? direction) => direction == null ? null :
            direction.Val == null || direction.Val.Value != OnOffOnlyValues.Off;
    }
}
