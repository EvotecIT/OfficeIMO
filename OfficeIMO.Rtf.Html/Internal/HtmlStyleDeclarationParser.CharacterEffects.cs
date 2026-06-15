namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlStyleDeclarationParser {
    private static bool? ParseVisibility(string value) {
        switch (value) {
            case "hidden":
            case "collapse":
                return true;
            case "visible":
                return false;
            default:
                return null;
        }
    }

    private static bool? ParseTextShadow(string value) => value == "none" ? false : true;

    private static bool? ParseBoolean(string value) {
        switch (value) {
            case "true":
            case "1":
            case "yes":
                return true;
            case "false":
            case "0":
            case "no":
            case "none":
                return false;
            default:
                return null;
        }
    }
}
