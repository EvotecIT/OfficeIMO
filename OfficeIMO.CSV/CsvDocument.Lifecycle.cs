using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument {
    /// <summary>Gets the path used by parameterless saves; stream and text loads have no associated path.</summary>
    public string? FilePath { get; private set; }

    /// <summary>Creates an empty in-memory CSV document.</summary>
    public static CsvDocument Create() => new CsvDocument();

    /// <summary>Creates an empty document associated with a path without writing a file.</summary>
    public static CsvDocument Create(string path) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return new CsvDocument { FilePath = Path.GetFullPath(path) };
    }

    /// <summary>Saves to the path of the last successful path load or save.</summary>
    public void Save(CsvSaveOptions? options = null) => Save(AssociatedPath(), options);

    /// <summary>Asynchronously saves to the associated path.</summary>
    public Task SaveAsync(CsvSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        SaveAsync(AssociatedPath(), options, cancellationToken);

    private string AssociatedPath() => FilePath ?? throw new InvalidOperationException("The document has no associated path. Save to a path first.");

    /// <summary>Completes this managed document's lifetime. No file or caller-owned stream is held open.</summary>
    /// <remarks>Existing row data remains available. Each streaming enumeration owns and closes its own reader.</remarks>
    public void Dispose() { }
}
