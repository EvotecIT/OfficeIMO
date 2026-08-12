namespace OfficeIMO.Pdf;

/// <summary>A named document-level JavaScript action from the catalog JavaScript name tree.</summary>
public sealed class PdfJavaScript {
    internal PdfJavaScript(string name, string script) {
        Name = name;
        Script = script;
    }

    /// <summary>Name-tree key that identifies the script.</summary>
    public string Name { get; }

    /// <summary>Exact JavaScript source stored by the PDF action.</summary>
    public string Script { get; }
}
