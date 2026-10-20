using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordMath {
        /// <summary>Identifies operator presentation that editable OMML cannot retain.</summary>
        /// <remarks>Raw import reports authored side lengths even outside the shared renderer's bounded grammar.</remarks>
        internal static bool HasUnsupportedNativeOperatorProperties(string? form, string? leftSpace, string? rightSpace) =>
            OfficeMathOperatorSpacing.NormalizeForm(form) != null
            || !string.IsNullOrWhiteSpace(leftSpace)
            || !string.IsNullOrWhiteSpace(rightSpace);
    }
}
