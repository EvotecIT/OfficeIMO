namespace OfficeIMO.Html.Runtime;

/// <summary>Versioned runtime behavior selected for a trusted document session.</summary>
public enum HtmlRuntimeProfile {
    /// <summary>A single live document with scripts, resources, storage, history routes, automation, waits, and independent capture.</summary>
    ScriptedDocumentV1,
    /// <summary>The scripted-document contract plus bounded cross-document navigation, reload, and history restoration.</summary>
    WebApplicationV1
}
