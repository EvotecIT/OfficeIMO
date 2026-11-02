using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Foundation;

/// <summary>Binds native workbook save options and the option factory's friendly hashtable keys.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeExcelSaveOptionsTransformationAttribute : ArgumentTransformationAttribute {
    private static readonly Dictionary<string, string> Names = new(StringComparer.OrdinalIgnoreCase) {
        ["Preflight"] = nameof(ExcelSaveOptions.SafePreflight),
        ["RepairDefinedNames"] = nameof(ExcelSaveOptions.SafeRepairDefinedNames),
        ["NoFastPackageWriter"] = nameof(ExcelSaveOptions.DisableFastPackageWriter),
        ["EvaluateFormulas"] = nameof(ExcelSaveOptions.EvaluateFormulasBeforeSave),
        ["ClearCachedFormulaResults"] = nameof(ExcelSaveOptions.ClearCachedFormulaResultsBeforeSave),
        ["MarkFormulasDirty"] = nameof(ExcelSaveOptions.MarkFormulasDirtyBeforeSave),
        ["FullCalculationOnOpen"] = nameof(ExcelSaveOptions.ForceFullCalculationOnOpen)
    };

    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value is IDictionary dictionary) {
            var normalized = new Hashtable(StringComparer.OrdinalIgnoreCase);
            foreach (DictionaryEntry entry in dictionary) {
                if (entry.Key is not string key) throw OfficeArgumentValue.Invalid("Save option names must be strings.");
                string name = Names.TryGetValue(key, out string? native) ? native : key;
                if (normalized.ContainsKey(name)) throw OfficeArgumentValue.Invalid($"Specify the save option '{name}' only once.");
                normalized.Add(name, entry.Value);
            }
            value = normalized;
        }
        return new OfficeOptionsTransformationAttribute(typeof(ExcelSaveOptions)).Transform(engineIntrinsics, value!);
    }
}
