using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Creates template-free native Numbers spreadsheets with positioned scalar tables.</summary>
/// <remarks>Instances are not thread safe. This creation model does not edit loaded packages or author formulas.</remarks>
public sealed class IWorkNumbersDocument {
    private readonly List<IWorkNumbersSheetBuilder> _sheets = new();

    private IWorkNumbersDocument() => Sheets = _sheets.AsReadOnly();

    /// <summary>Gets sheets in insertion order.</summary>
    public IReadOnlyList<IWorkNumbersSheetBuilder> Sheets { get; }

    /// <summary>Creates an empty spreadsheet. Add at least one sheet and a table to each sheet before saving.</summary>
    public static IWorkNumbersDocument Create() => new();

    /// <summary>Adds a sheet with a unique name, compared without case.</summary>
    public IWorkNumbersSheetBuilder AddSheet(string name) {
        IWorkCreationValidation.ValidateName(name, nameof(name));
        if (_sheets.Any(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("Sheet names must be unique without regard to case.", nameof(name));
        var sheet = new IWorkNumbersSheetBuilder(name); _sheets.Add(sheet); return sheet;
    }

    /// <summary>Encodes a deterministic native Numbers package without a template or native application.</summary>
    public byte[] SaveBytes(IWorkNumbersWriteOptions? options = null, CancellationToken cancellationToken = default) =>
        IWorkNumbersPackageWriter.Write(this, (options ?? new IWorkNumbersWriteOptions()).Snapshot(), cancellationToken);

    /// <summary>Writes a completed package at the stream's current position and leaves the stream open.</summary>
    /// <remarks>Encoding finishes before copying. A copy failure can leave partial stream data.</remarks>
    public void Save(Stream destination, IWorkNumbersWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        IWorkNativePackageWriter.Copy(SaveBytes(options, cancellationToken), destination, cancellationToken);
    }

    /// <summary>Atomically saves a native package, rejecting an existing file unless overwrite is enabled.</summary>
    /// <remarks>Validation and encoding finish before staging; failed staging preserves the existing destination.</remarks>
    public void Save(string path, bool overwrite = false, IWorkNumbersWriteOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A destination path is required.", nameof(path));
        byte[] bytes = SaveBytes(options, cancellationToken);
        OfficeFileCommit.WriteAtomically(path, stream => IWorkNativePackageWriter.Copy(bytes, stream, cancellationToken),
            cancellationToken, overwrite ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
    }
}
