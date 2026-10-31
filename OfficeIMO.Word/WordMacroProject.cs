using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

/// <summary>A VBA project attached to a Word document. Reading it never executes macro code.</summary>
public sealed class WordMacroProject {
    private readonly WordDocument _document;
    private readonly VbaProjectPart _part;

    internal WordMacroProject(WordDocument document, VbaProjectPart part) {
        _document = document;
        _part = part;
    }

    /// <summary>Gets the owning document.</summary>
    public WordDocument Document => _document;

    /// <summary>Gets the attached project's package URI.</summary>
    public string PartUri { get { EnsureAttached(); return _part.Uri.ToString(); } }

    /// <summary>Gets the project's standard, class and host modules.</summary>
    public IReadOnlyList<WordMacro> Modules { get { EnsureAttached(); return _document.Macros; } }

    /// <summary>Reads the attached project bytes subject to the specified maximum size.</summary>
    public byte[] ToBytes(long maximumBytes = OfficeVbaProjectInfo.DefaultMaximumProjectBytes) {
        EnsureAttached();
        return _document.ExtractMacros(maximumBytes);
    }

    /// <summary>Saves the attached project bytes to a file.</summary>
    public void Save(string path, long maximumBytes = OfficeVbaProjectInfo.DefaultMaximumProjectBytes) {
        EnsureAttached();
        _document.SaveMacros(path, maximumBytes);
    }

    /// <summary>Removes this project from its document.</summary>
    public void Remove() {
        EnsureAttached();
        _document.RemoveMacros();
    }

    private void EnsureAttached() {
        if (!ReferenceEquals(_document._wordprocessingDocument.MainDocumentPart?.VbaProjectPart, _part))
            throw new InvalidOperationException("The VBA project has been removed or replaced.");
    }
}
