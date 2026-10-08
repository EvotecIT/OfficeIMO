using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
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
            if (project.Modules.Any(module => module.Kind != OfficeVbaModuleKind.Document
                && (sheetNames.Contains(module.Name, StringComparer.OrdinalIgnoreCase)
                    || string.Equals(module.Name, workbookName, StringComparison.OrdinalIgnoreCase)))) {
                throw new ArgumentException("A non-document VBA module name collides with an existing Excel host code name.", nameof(project));
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
            OfficeVbaProjectPartEditor.EnsureCanApply(workbookPart.VbaProjectPart, bytes, options);
            SpreadsheetDocumentType originalType = _spreadSheetDocument.DocumentType;
            StringValue? originalCodeName = properties?.CodeName;
            bool addedCodeName = workbookModules.Length == 1 && string.IsNullOrEmpty(existingName);
            bool hadProperties = properties != null;
            try {
                if (workbookModules.Length == 1 && string.IsNullOrEmpty(existingName)) {
                    if (properties == null) { properties = new WorkbookProperties(); workbook.AddChild(properties, true); }
                    properties.CodeName = workbookModules[0].Name;
                }
                // Convert the carrier before replacing VBA, then use the resulting owner.
                // ChangeDocumentType can replace WorkbookPart and its attached part objects.
                EnsureMacroEnabledDocumentType();
                WorkbookPart current = WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is missing.");
                bool changed = OfficeVbaProjectPartEditor.Apply(current, current.VbaProjectPart, bytes, options, out _);
                if (changed || addedCodeName || originalType != _spreadSheetDocument.DocumentType) MarkPackageDirty();
            } catch (Exception failure) {
                try {
                    WorkbookPart current = WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is missing during recovery.");
                    Workbook root = current.Workbook ?? throw new InvalidOperationException("Workbook root is missing during recovery.");
                    WorkbookProperties? currentProperties = root.GetFirstChild<WorkbookProperties>();
                    if (addedCodeName && currentProperties != null) {
                        if (hadProperties) currentProperties.CodeName = originalCodeName;
                        else root.RemoveChild(currentProperties);
                    }
                    if (_spreadSheetDocument.DocumentType != originalType) {
                        root.Save(); _spreadSheetDocument.ChangeDocumentType(originalType);
                        _workBookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing after recovery.");
                    }
                } catch (Exception restorationFailure) {
                    throw new AggregateException("The VBA update failed and Excel host metadata could not be restored. Discard this document instance.", failure, restorationFailure);
                }
                throw;
            }
        });
    }
}
