using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for importing worksheets from one workbook into another.
    /// </summary>
    public sealed class ExcelWorkbookMergeOptions {
        /// <summary>Creates an independent copy of the workbook import settings, selected sheet names and target sheet names.</summary>
        public ExcelWorkbookMergeOptions Clone() => new ExcelWorkbookMergeOptions {
            SheetNames = SheetNames?.ToArray(),
            SheetNamePrefix = SheetNamePrefix,
            TargetSheetNames = TargetSheetNames is Dictionary<string, string> targetSheetNames
                ? new Dictionary<string, string>(targetSheetNames, targetSheetNames.Comparer)
                : TargetSheetNames?.ToDictionary(_ => _.Key, _ => _.Value, StringComparer.Ordinal),
            SheetNameValidationMode = SheetNameValidationMode,
            CopyMode = CopyMode,
            CopyExternalWorkbookReferences = CopyExternalWorkbookReferences,
            MaxDefinedNames = MaxDefinedNames,
            MaxDefinedNameCharacters = MaxDefinedNameCharacters,
            MaxInCellImages = MaxInCellImages,
            MaxInCellImageBytes = MaxInCellImageBytes,
            MaxTotalInCellImageBytes = MaxTotalInCellImageBytes
        };

        /// <summary>
        /// Gets or sets source worksheet names to import. When empty, all source worksheets are imported.
        /// </summary>
        public IReadOnlyList<string>? SheetNames { get; set; }

        /// <summary>
        /// Gets or sets an optional prefix added to every imported worksheet name.
        /// </summary>
        public string? SheetNamePrefix { get; set; }

        /// <summary>
        /// Gets or sets explicit target worksheet names keyed by source worksheet name. Keys match source
        /// worksheet names case-insensitively and must name imported worksheets. A worksheet with an entry takes
        /// that name as given, without <see cref="SheetNamePrefix"/>; a worksheet without one keeps the prefixed
        /// source name. Names are validated with <see cref="SheetNameValidationMode"/>. When this map is nonempty and
        /// <see cref="ExcelSheetNameValidationMode.Strict"/> is selected, every requested name is checked before any worksheet is imported.
        /// References between imported worksheets map each source name directly to its target name, so a target
        /// name may reuse the name of another source worksheet.
        /// </summary>
        public IReadOnlyDictionary<string, string>? TargetSheetNames { get; set; }

        /// <summary>
        /// Gets or sets sheet name validation behavior for imported worksheets.
        /// </summary>
        public ExcelSheetNameValidationMode SheetNameValidationMode { get; set; } = ExcelSheetNameValidationMode.Sanitize;

        /// <summary>
        /// Gets or sets the worksheet copy strategy used for imported sheets.
        /// </summary>
        public ExcelWorksheetCopyMode CopyMode { get; set; } = ExcelWorksheetCopyMode.Values;

        /// <summary>
        /// Gets or sets whether package-mode imports may preserve external-workbook references.
        /// External links are rejected unless callers opt in explicitly.
        /// </summary>
        public bool CopyExternalWorkbookReferences { get; set; }

        /// <summary>
        /// Gets or sets the maximum number of referenced defined names copied in package mode.
        /// </summary>
        public int MaxDefinedNames { get; set; } = 4096;

        /// <summary>
        /// Gets or sets the maximum aggregate formula characters across copied defined names.
        /// </summary>
        public int MaxDefinedNameCharacters { get; set; } = 1_000_000;

        /// <summary>
        /// Gets or sets the maximum number of in-cell image references copied across package-mode imports.
        /// </summary>
        public int MaxInCellImages { get; set; } = 1024;

        /// <summary>
        /// Gets or sets the maximum payload size of one in-cell image copied in package mode.
        /// </summary>
        public long MaxInCellImageBytes { get; set; } = 32_000_000;

        /// <summary>
        /// Gets or sets the maximum aggregate resolved in-cell image bytes copied across package-mode imports.
        /// Shared image assets are charged once per referencing cell.
        /// </summary>
        public long MaxTotalInCellImageBytes { get; set; } = 64_000_000;
    }
}
