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
            var sheetNames = workbookPart.WorksheetParts.Select(part => part.Worksheet?.SheetProperties?.CodeName?.Value)
                .Concat(workbookPart.ChartsheetParts.Select(part => part.Chartsheet?.ChartSheetProperties?.CodeName?.Value))
                .Where(name => !string.IsNullOrEmpty(name)).ToArray();
            string? workbookName = workbook.GetFirstChild<WorkbookProperties>()?.CodeName?.Value;
            if (sheetNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() != sheetNames.Length
                || !string.IsNullOrEmpty(workbookName) && sheetNames.Contains(workbookName, StringComparer.OrdinalIgnoreCase)) {
                throw new ArgumentException("Excel workbook and sheet code names must be unique.", nameof(project));
            }
            foreach (OfficeVbaModule module in project.Modules.Where(module => module.Kind == OfficeVbaModuleKind.Document)) {
                string? identity = OfficeIMO.Core.Internal.OfficeVbaText.GetBaseIdentity(module.Source);
                if (string.Equals(identity, "0{00020819-0000-0000-C000-000000000046}", StringComparison.OrdinalIgnoreCase)) {
                    if (sheetNames.Contains(module.Name, StringComparer.OrdinalIgnoreCase)) {
                        throw new ArgumentException("The VBA workbook module name collides with an existing sheet code name.", nameof(project));
                    }
                    continue;
                }
                int matchingSheets;
                if (string.Equals(identity, "0{00020820-0000-0000-C000-000000000046}", StringComparison.OrdinalIgnoreCase)) {
                    matchingSheets = workbookPart.WorksheetParts.Count(part => string.Equals(
                        part.Worksheet?.SheetProperties?.CodeName?.Value, module.Name, StringComparison.OrdinalIgnoreCase));
                } else if (string.Equals(identity, "0{00020821-0000-0000-C000-000000000046}", StringComparison.OrdinalIgnoreCase)) {
                    matchingSheets = workbookPart.ChartsheetParts.Count(part => string.Equals(
                        part.Chartsheet?.ChartSheetProperties?.CodeName?.Value, module.Name, StringComparison.OrdinalIgnoreCase));
                } else {
                    throw new ArgumentException("The VBA document module has no supported Excel host identity.", nameof(project));
                }
                if (matchingSheets != 1) throw new ArgumentException("The VBA sheet module must match exactly one existing sheet code name and type.", nameof(project));
            }
            OfficeVbaModule[] workbookModules = project.Modules.Where(module => module.Kind == OfficeVbaModuleKind.Document
                && string.Equals(OfficeIMO.Core.Internal.OfficeVbaText.GetBaseIdentity(module.Source),
                    "0{00020819-0000-0000-C000-000000000046}", StringComparison.OrdinalIgnoreCase)).ToArray();
            if (workbookModules.Length > 1) throw new ArgumentException("An Excel project cannot contain multiple workbook document modules.", nameof(project));
            WorkbookProperties? properties = workbook.GetFirstChild<WorkbookProperties>();
            string? existingName = properties?.CodeName?.Value;
            if (workbookModules.Length == 1 && !string.IsNullOrEmpty(existingName)
                && !string.Equals(existingName, workbookModules[0].Name, StringComparison.OrdinalIgnoreCase)) {
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
