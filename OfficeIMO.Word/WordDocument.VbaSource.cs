using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Reads a detached, editable native VBA project; returns null when no project exists.</summary>
    public OfficeVbaProject? ReadVbaProject(OfficeVbaReadOptions? options = null) {
        options ??= new OfficeVbaReadOptions();
        VbaProjectPart? part = _wordprocessingDocument.MainDocumentPart?.VbaProjectPart;
        return part == null ? null : OfficeVbaProject.Load(OfficeVbaProjectPartEditor.Read(part, options.MaximumProjectBytes), options);
    }

    /// <summary>Applies a native source project while preserving Word's VBA data and requiring explicit signature removal.</summary>
    public void SetVbaProject(OfficeVbaProject project, OfficeVbaWriteOptions? options = null) {
        if (project == null) throw new ArgumentNullException(nameof(project));
        options ??= new OfficeVbaWriteOptions();
        if (project.Modules.Any(module => module.Kind == OfficeVbaModuleKind.Document
            && OfficeIMO.Core.Internal.OfficeVbaText.GetBaseIdentity(module.Source)?.StartsWith("0{", StringComparison.Ordinal) == true)) {
            throw new ArgumentException("A document module from a different Office host cannot be applied to Word.", nameof(project));
        }
        byte[] bytes = project.Write(options).GetBytes();
        if (!project.Modules.Any(module => module.Kind == OfficeVbaModuleKind.Document)) {
            // A newly authored host-neutral project needs Word's document class to be loadable.
            // Prepare a detached copy so applying it does not mutate the caller's project model.
            OfficeVbaProject prepared = OfficeVbaProject.Load(bytes, new OfficeVbaReadOptions {
                MaximumProjectBytes = options.MaximumProjectBytes, MaximumExpandedBytes = options.MaximumExpandedBytes
            });
            prepared.AddDocumentModuleWithIdentity("ThisDocument", "Attribute VB_TemplateDerived = True\r\n", "1Normal.ThisDocument");
            bytes = prepared.Write(options).GetBytes();
        }
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing.");
        if (OfficeVbaProjectPartEditor.Apply(main, main.VbaProjectPart, bytes, options, out VbaProjectPart part) && part.VbaDataPart == null) {
            VbaDataPart data = part.AddNewPart<VbaDataPart>();
            data.VbaSuppData = new DocumentFormat.OpenXml.Office.Word.VbaSuppData();
        }
    }
}
