using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Reads a detached, editable native VBA project; returns null when no project exists.</summary>
    public OfficeVbaProject? ReadVbaProject(OfficeVbaReadOptions? options = null) {
        options ??= new OfficeVbaReadOptions();
        return Locking.ExecuteRead(EnsureLock(), () => {
            VbaProjectPart? part = WorkbookPartRoot?.VbaProjectPart;
            return part == null ? null : OfficeVbaProject.Load(OfficeVbaProjectPartEditor.Read(part, options.MaximumProjectBytes), options);
        });
    }

    /// <summary>Applies a native source project, preserving unrelated payloads and requiring explicit VBA signature removal.</summary>
    public void SetVbaProject(OfficeVbaProject project, OfficeVbaWriteOptions? options = null) {
        if (project == null) throw new ArgumentNullException(nameof(project));
        options ??= new OfficeVbaWriteOptions();
        byte[] bytes = project.Write(options).GetBytes();
        Locking.ExecuteWrite(EnsureLock(), () => {
            WorkbookPart workbookPart = WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is missing.");
            Workbook workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook root is missing.");
            OfficeVbaModule[] workbookModules = project.Modules.Where(module => module.Kind == OfficeVbaModuleKind.Document
                && string.Equals(OfficeIMO.Core.Internal.OfficeVbaText.GetBaseIdentity(module.Source),
                    "0{00020819-0000-0000-C000-000000000046}", StringComparison.OrdinalIgnoreCase)).ToArray();
            if (workbookModules.Length > 1) throw new ArgumentException("An Excel project cannot contain multiple workbook document modules.", nameof(project));
            WorkbookProperties? properties = workbook.GetFirstChild<WorkbookProperties>();
            string? existingName = properties?.CodeName?.Value;
            if (workbookModules.Length == 1 && !string.IsNullOrEmpty(existingName)
                && !existingName.Equals(workbookModules[0].Name, StringComparison.OrdinalIgnoreCase)) {
                throw new ArgumentException("The VBA workbook module does not match the document's existing code name.", nameof(project));
            }
            VbaProjectPart part = workbookPart.VbaProjectPart ?? workbookPart.AddNewPart<VbaProjectPart>();
            if (OfficeVbaProjectPartEditor.Apply(part, bytes, options)) {
                if (workbookModules.Length == 1 && string.IsNullOrEmpty(existingName)) {
                    if (properties == null) { properties = new WorkbookProperties(); workbook.AddChild(properties, true); }
                    properties.CodeName = workbookModules[0].Name;
                }
                EnsureMacroEnabledDocumentType();
                MarkPackageDirty();
            }
        });
    }
}
