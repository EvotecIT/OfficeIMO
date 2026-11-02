using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel;

/// <summary>A VBA project attached to an Excel workbook. Reading it never executes macro code.</summary>
public sealed class ExcelMacroProject {
    private readonly VbaProjectPart _part;

    internal ExcelMacroProject(ExcelDocument document, VbaProjectPart part) {
        Document = document;
        _part = part;
    }

    /// <summary>Gets the owning workbook.</summary>
    public ExcelDocument Document { get; }

    /// <summary>Gets the attached project's package URI.</summary>
    public string PartUri => Locking.ExecuteRead(Document.EnsureLock(), () => {
        EnsureAttached();
        return _part.Uri.ToString();
    });

    /// <summary>Reads the attached project's bytes subject to the specified maximum size.</summary>
    public byte[] ToBytes(long maximumBytes = OfficeVbaProjectInfo.DefaultMaximumProjectBytes) => Locking.ExecuteRead(Document.EnsureLock(), () => {
        EnsureAttached();
        return Document.ExtractMacros(maximumBytes);
    });

    /// <summary>Saves the attached project bytes to a file.</summary>
    public void Save(string path, long maximumBytes = OfficeVbaProjectInfo.DefaultMaximumProjectBytes) => Locking.ExecuteRead(Document.EnsureLock(), () => {
        EnsureAttached();
        Document.SaveMacros(path, maximumBytes);
    });

    /// <summary>Removes this project from its workbook.</summary>
    public void Remove() => Locking.ExecuteWrite(Document.EnsureLock(), () => {
        EnsureAttached();
        Document.RemoveMacros();
    });

    private void EnsureAttached() {
        if (!ReferenceEquals(Document.WorkbookPartRoot?.VbaProjectPart, _part))
            throw new InvalidOperationException("The VBA project has been removed or replaced.");
    }
}
