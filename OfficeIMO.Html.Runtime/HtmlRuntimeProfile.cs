namespace OfficeIMO.Html.Runtime;

/// <summary>Versioned runtime behavior selected for a trusted document session.</summary>
public enum HtmlRuntimeProfile {
    /// <summary>A root live document with qualified same-origin child realms, scripts, resources, storage, history routes, automation, waits, and independent capture.</summary>
    ScriptedDocumentV1,
    /// <summary>The scripted-document contract plus bounded cross-document navigation, reload, and history restoration.</summary>
    WebApplicationV1
}
