using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Reads a detached, editable native VBA project; returns null when no project exists.</summary>
    public OfficeVbaProject? ReadVbaProject(OfficeVbaReadOptions? options = null) {
        ThrowIfDisposed();
        options ??= new OfficeVbaReadOptions();
        var part = _presentationPart.VbaProjectPart;
        return part == null ? null : OfficeVbaProject.Load(OfficeVbaProjectPartEditor.Read(part, options.MaximumProjectBytes), options);
    }

    /// <summary>Applies a native source project, preserving non-signature child parts and requiring explicit signature removal.</summary>
    public void SetVbaProject(OfficeVbaProject project, OfficeVbaWriteOptions? options = null) {
        ThrowIfDisposed();
        if (project == null) throw new ArgumentNullException(nameof(project));
        options ??= new OfficeVbaWriteOptions();
        if (project.Modules.Any(module => module.Kind == OfficeVbaModuleKind.Document)) {
            throw new ArgumentException("PowerPoint source projects cannot contain document modules bound to another Office host.", nameof(project));
        }
        byte[] bytes = project.Write(options).GetBytes();
        OfficeVbaProjectPartEditor.Apply(_presentationPart, _presentationPart.VbaProjectPart, bytes, options, out _);
    }
}
