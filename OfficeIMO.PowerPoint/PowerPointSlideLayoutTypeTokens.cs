namespace OfficeIMO.PowerPoint {
    /// <summary>The <c>p:sldLayout/@type</c> tokens of <see cref="PowerPointSlideLayoutType"/>.</summary>
    internal static class PowerPointSlideLayoutTypeTokens {
        private static readonly string[] Tokens = {
            "title", "tx", "twoColTx", "tbl", "txAndChart", "chartAndTx", "dgm", "chart", "txAndClipArt", "clipArtAndTx", "titleOnly",
            "blank", "txAndObj", "objAndTx", "objOnly", "obj", "txAndMedia", "mediaAndTx", "objOverTx", "txOverObj", "txAndTwoObj",
            "twoObjAndTx", "twoObjOverTx", "fourObj", "vertTx", "clipArtAndVertTx", "vertTitleAndTx", "vertTitleAndTxOverChart",
            "twoObj", "objAndTwoObj", "twoObjAndObj", "cust", "secHead", "twoTxTwoObj", "objTx", "picTx"
        };

        internal static string ToOpenXmlToken(this PowerPointSlideLayoutType value) {
            int index = (int)value;
            if (index < 0 || index >= Tokens.Length) throw new ArgumentOutOfRangeException(nameof(value), value, "Unsupported PowerPoint slide layout type.");
            return Tokens[index];
        }

        internal static bool TryParse(string? token, out PowerPointSlideLayoutType value) {
            int index = token == null ? -1 : Array.IndexOf(Tokens, token);
            value = (PowerPointSlideLayoutType)Math.Max(index, 0);
            return index >= 0;
        }
    }
}
