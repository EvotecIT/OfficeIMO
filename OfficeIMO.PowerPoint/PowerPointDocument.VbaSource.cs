using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    /// <summary>Reads a detached, editable native VBA project; returns null when no project exists.</summary>
    public OfficeVbaProject? ReadVbaProject(OfficeVbaReadOptions? options = null) {
        ThrowIfDisposed();
        options ??= new OfficeVbaReadOptions();
        if (options.MaximumProjectBytes <= 0 || !TryReadPackageVbaProject(options.MaximumProjectBytes, inspectSignature: false,
            out _, out byte[]? bytes, validateForSource: true)) bytes = ReadSourceVbaBytesUsingSdk(options.MaximumProjectBytes);
        return bytes == null ? null : OfficeVbaProject.Load(bytes, options);
    }

    /// <summary>Applies a native source project, preserving non-signature child parts and requiring explicit signature removal.</summary>
    public void SetVbaProject(OfficeVbaProject project, OfficeVbaWriteOptions? options = null) {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(project);
        options ??= new OfficeVbaWriteOptions();
        if (project.Modules.Any(module => module.Kind == OfficeVbaModuleKind.Document)) {
            throw new ArgumentException("PowerPoint source projects cannot contain document modules bound to another Office host.", nameof(project));
        }
        byte[] bytes = project.Write(options).GetBytes();
        if (!TryMutatePackageVbaProject(bytes, out _, options)) SetSourceVbaProjectUsingSdk(bytes, options);
    }

    private static bool PackageVbaBytesMatch(OfficeOpcEditablePackage package, string part, byte[] bytes) {
        if (package.TryGetBinaryLength(part, out long length) && length != bytes.LongLength) return false;
        using Stream input = package.OpenPartContentStream(part, default);
        try {
            return OfficeStreamReader.ReadAllBytes(input, Math.Max(1L, bytes.LongLength)).SequenceEqual(bytes);
        } catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
            return false;
        }
    }

    private static void ValidatePackageVbaRecovery(OfficeOpcEditablePackage package, IReadOnlyCollection<string> parts, long remaining) {
        if (parts.Count > 1024) throw new InvalidDataException("The VBA recovery subgraph exceeds the part limit.");
        foreach (string part in parts) {
            using Stream input = package.OpenPartContentStream(part, default);
            remaining -= OfficeStreamReader.ReadAllBytes(input, Math.Max(1L, remaining)).LongLength;
            if (remaining < 0) throw new InvalidDataException("The VBA recovery subgraph exceeds the configured recovery byte limit.");
        }
    }
}
