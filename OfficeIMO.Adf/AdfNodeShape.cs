namespace OfficeIMO.Adf;

/// <summary>Common native shape rules shared by JSON writing and structural validation.</summary>
internal static class AdfNodeShape {
    internal static bool RequiresContent(string type) => type is
        "blockquote" or "bulletList" or "orderedList" or "listItem" or "taskList" or
        "table" or "tableRow" or "tableCell" or "tableHeader" or "mediaSingle" or
        "mediaGroup" or "panel" or "bodiedExtension";

    internal static int MinimumChildren(string type) => RequiresContent(type) && type != "tableRow" ? 1 : 0;

    internal static bool AllowsMark(string nodeType, string markType) => nodeType switch {
        "text" => markType is "strong" or "em" or "code" or "strike" or "underline" or "link" or
            "subsup" or "textColor" or "backgroundColor" or "annotation",
        "paragraph" => markType is "alignment" or "indentation" or "fontSize",
        "heading" => markType is "alignment" or "indentation",
        "mediaSingle" => markType == "link",
        "media" => markType is "link" or "annotation" or "border" or "dataConsumer",
        "table" => markType == "fragment",
        "extension" or "inlineExtension" or "bodiedExtension" => true,
        _ => false
    };
}
