using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Word.Html;

public static partial class WordHtmlConverterExtensions {
    internal static void NormalizeRoleTables(
        IHtmlDocument document,
        HtmlDiagnosticReport? diagnostics,
        Action<IElement> materializeCss,
        ISet<IElement> materializedElements,
        Action<IElement, IElement> registerNative) =>
        HtmlRoleTableNormalizer.Normalize(
            document,
            registerNative,
            materializeCss,
            materializedElements,
            table => diagnostics?.Add(
                "OfficeIMO.Word.Html",
                HtmlConversionDiagnosticCodes.ContentApproximated,
                "An ARIA table with unsupported row or cell structure remains in document flow rather than becoming an editable Word table.",
                HtmlDiagnosticSeverity.Warning,
                "role=table",
                "unsupported ARIA table structure",
                OfficeConversionLossKind.Approximation));
}
