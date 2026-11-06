using System.Threading;

namespace OfficeIMO.IWork;

public sealed partial class IWorkPagesDocument : IDisposable {
    /// <summary>Gets the associated output path; stream saves do not change it.</summary>
    public string? FilePath { get; private set; }

    /// <summary>Creates a native authoring model associated with a path, without writing a file.</summary>
    public static IWorkPagesDocument Create(string path) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var document = Create(); document.FilePath = Path.GetFullPath(path); return document;
    }

    /// <summary>Saves to the associated path; replacing an existing file still requires explicit overwrite.</summary>
    public void Save(bool overwrite = false, IWorkPagesSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        Save(FilePath ?? throw new InvalidOperationException("The document has no associated path. Save to a path first."), overwrite, options, cancellationToken);

    /// <summary>Completes this managed authoring model's lifetime. No native application, file or caller stream is retained.</summary>
    public void Dispose() { }
}
