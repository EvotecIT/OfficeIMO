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
        byte[] bytes = project.Write(options).GetBytes();
        var part = _presentationPart.VbaProjectPart ?? _presentationPart.AddNewPart<DocumentFormat.OpenXml.Packaging.VbaProjectPart>();
        OfficeVbaProjectPartEditor.Apply(part, bytes, options);
    }
}
