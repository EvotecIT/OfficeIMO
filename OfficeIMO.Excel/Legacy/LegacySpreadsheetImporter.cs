using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Legacy;

/// <summary>Detects and imports selected legacy spreadsheets without executing macros or resolving external links.</summary>
public static class LegacySpreadsheetImporter {
    private static readonly ILegacySpreadsheetAdapter[] Adapters = {
        new QuattroProAdapter(),
        new MicrosoftWorksSpreadsheetAdapter(),
        new MultiplanAdapter(),
        new Lotus123Adapter()
    };

    /// <summary>Detects a legacy spreadsheet file.</summary>
    public static LegacySpreadsheetDetection Detect(string path, LegacySpreadsheetImportOptions? options = null, CancellationToken cancellationToken = default) {
        LegacySpreadsheetImportOptions effective = Prepare(options, path);
        return Detect(OfficeLegacyImportBuffer.ReadAll(path, effective.Limits, cancellationToken), effective, cancellationToken);
    }

    /// <summary>Detects a legacy spreadsheet stream.</summary>
    public static LegacySpreadsheetDetection Detect(Stream stream, LegacySpreadsheetImportOptions? options = null, CancellationToken cancellationToken = default) {
        LegacySpreadsheetImportOptions effective = Prepare(options, options?.SourceName);
        return Detect(OfficeLegacyImportBuffer.ReadAll(stream, effective.Limits, cancellationToken), effective, cancellationToken);
    }

    /// <summary>Detects legacy spreadsheet bytes.</summary>
    public static LegacySpreadsheetDetection Detect(byte[] data, LegacySpreadsheetImportOptions? options = null, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        LegacySpreadsheetImportOptions effective = Prepare(options, options?.SourceName);
        if (data.Length > effective.Limits.MaxInputBytes) throw new InvalidDataException("Legacy spreadsheet input exceeds the configured byte limit.");
        cancellationToken.ThrowIfCancellationRequested();
        return SelectAdapter(data, effective, cancellationToken).Detection;
    }

    /// <summary>Imports a legacy spreadsheet file into a normal editable <see cref="ExcelDocument"/>.</summary>
    public static LegacySpreadsheetImportResult Import(string path, LegacySpreadsheetImportOptions? options = null, CancellationToken cancellationToken = default) {
        LegacySpreadsheetImportOptions effective = Prepare(options, path);
        return Import(OfficeLegacyImportBuffer.ReadAll(path, effective.Limits, cancellationToken), effective, cancellationToken);
    }

    /// <summary>Imports a legacy spreadsheet stream into a normal editable <see cref="ExcelDocument"/>.</summary>
    public static LegacySpreadsheetImportResult Import(Stream stream, LegacySpreadsheetImportOptions? options = null, CancellationToken cancellationToken = default) {
        LegacySpreadsheetImportOptions effective = Prepare(options, options?.SourceName);
        return Import(OfficeLegacyImportBuffer.ReadAll(stream, effective.Limits, cancellationToken), effective, cancellationToken);
    }

    /// <summary>Imports legacy spreadsheet bytes into a normal editable <see cref="ExcelDocument"/>.</summary>
    public static LegacySpreadsheetImportResult Import(byte[] data, LegacySpreadsheetImportOptions? options = null, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        LegacySpreadsheetImportOptions effective = Prepare(options, options?.SourceName);
        if (data.Length > effective.Limits.MaxInputBytes) throw new InvalidDataException("Legacy spreadsheet input exceeds the configured byte limit.");
        (ILegacySpreadsheetAdapter adapter, LegacySpreadsheetDetection detection) = SelectAdapter(data, effective, cancellationToken);
        LegacySpreadsheetModel model = adapter.Parse(data, effective.Limits, cancellationToken);
        if (effective.RequireStructured && model.Quality != OfficeLegacyImportQuality.Structured) {
            throw new InvalidDataException($"The {detection.ProfileId} adapter produced salvage quality while structured import was required.");
        }

        PrepareNameProjection(model, cancellationToken);
        var report = new OfficeLegacyImportReport(detection.ProfileId, model.Quality, model.Findings, model.InertContent, model.RecoveredCellCount);
        var metadata = new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(model.Metadata, StringComparer.OrdinalIgnoreCase));
        var charts = Array.AsReadOnly(model.Charts.ToArray());
        var cells = Array.AsReadOnly(model.Sheets.SelectMany(static sheet => sheet.Cells.Select(cell => new LegacySpreadsheetCellContent(sheet.Name, cell))).ToArray());
        var names = Array.AsReadOnly(model.Names.Select(static name => new LegacySpreadsheetNameContent(name)).ToArray());
        ExcelDocument? document = null;
        try {
            document = Project(model, cancellationToken);
            return new LegacySpreadsheetImportResult(document, detection, report, metadata, charts, cells, names);
        } catch {
            document?.Dispose();
            throw;
        }
    }

    private static ExcelDocument Project(LegacySpreadsheetModel model, CancellationToken cancellationToken) {
        ExcelDocument document = ExcelDocument.Create();
        try {
            foreach (LegacySpreadsheetSheet sourceSheet in model.Sheets) {
                cancellationToken.ThrowIfCancellationRequested();
                ExcelSheet target = document.AddWorksheet(sourceSheet.Name);
                foreach (LegacySpreadsheetCell cell in sourceSheet.Cells) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.IsNullOrWhiteSpace(cell.Formula)) {
                        target.CellValue(cell.Row, cell.Column, cell.Value);
                        target.CellFormula(cell.Row, cell.Column, cell.Formula!);
                    } else if (cell.Value != null) target.CellValue(cell.Row, cell.Column, cell.Value);
                    else target.CellAt(cell.Row, cell.Column);
                    if (cell.Alignment.HasValue) target.CellAlign(cell.Row, cell.Column, cell.Alignment.Value);
                    if (!string.IsNullOrWhiteSpace(cell.NumberFormat)) target.CellAt(cell.Row, cell.Column).SetNumberFormat(cell.NumberFormat!);
                    if (!string.IsNullOrWhiteSpace(cell.Comment)) target.SetComment(cell.Row, cell.Column, cell.Comment!, "Legacy source");
                }
            }
            foreach (LegacySpreadsheetName name in model.Names) {
                cancellationToken.ThrowIfCancellationRequested();
                if (name.ProjectedName == null) continue;
                string reference = "'" + name.SheetName.Replace("'", "''") + "'!$" + ColumnName(name.FirstColumn) + "$" + name.FirstRow.ToString(CultureInfo.InvariantCulture) + ":$" + ColumnName(name.LastColumn) + "$" + name.LastRow.ToString(CultureInfo.InvariantCulture);
                document.SetNamedRange(name.ProjectedName, reference, save: false, validationMode: ExcelDefinedNameValidationMode.Strict);
            }
            if (model.Sheets.Count == 0) document.AddWorksheet("Sheet1");
            return document;
        } catch {
            document.Dispose();
            throw;
        }
    }

    private static void PrepareNameProjection(LegacySpreadsheetModel model, CancellationToken cancellationToken) {
        var projected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int collisionCount = 0;
        int invalidCount = 0;
        foreach (LegacySpreadsheetName name in model.Names) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                string validated = ExcelDocument.NormalizeDefinedName(name.Name, ExcelDefinedNameValidationMode.Strict);
                if (!projected.Add(validated)) {
                    model.Metadata["UnprojectedName.Collision." + model.Metadata.Count.ToString(CultureInfo.InvariantCulture)] = name.Name;
                    collisionCount++;
                    continue;
                }
                name.ProjectedName = validated;
            } catch (ArgumentException) {
                model.Metadata["UnprojectedName.Invalid." + model.Metadata.Count.ToString(CultureInfo.InvariantCulture)] = name.Name;
                invalidCount++;
            }
        }
        if (collisionCount > 0) {
            model.Metadata["UnprojectedNameCollisionCount"] = collisionCount.ToString(CultureInfo.InvariantCulture);
            model.Findings.Add(LegacySpreadsheetAdapterBase.LossFinding("WK_NAME_COLLISION", "Name", "One or more source names collided case-insensitively and were retained in the semantic snapshot without overwriting projected names; the total is available in metadata."));
        }
        if (invalidCount > 0) {
            model.Metadata["UnprojectedInvalidNameCount"] = invalidCount.ToString(CultureInfo.InvariantCulture);
            model.Findings.Add(LegacySpreadsheetAdapterBase.LossFinding("WK_NAME_INVALID", "Name", "One or more source names were invalid Excel defined names and were retained in the semantic snapshot without silent sanitization; the total is available in metadata."));
        }
    }

    private static string ColumnName(int column) {
        string result = string.Empty;
        while (column > 0) { column--; result = (char)('A' + column % 26) + result; column /= 26; }
        return result;
    }

    private static (ILegacySpreadsheetAdapter Adapter, LegacySpreadsheetDetection Detection) SelectAdapter(byte[] data, LegacySpreadsheetImportOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.FormatHint.HasValue) {
            ILegacySpreadsheetAdapter hinted = Adapters.Single(adapter => adapter.Format == options.FormatHint.Value);
            int confidence = hinted.Probe(data, options.SourceName, options.Limits, cancellationToken, out string evidence);
            return (hinted, new LegacySpreadsheetDetection(hinted.Format, hinted.GetProfileId(data, options.Limits, cancellationToken), Math.Max(1, confidence),
                confidence == 0 ? "Explicit caller format hint." : evidence + " Explicit caller format hint confirmed the family."));
        }
        ILegacySpreadsheetAdapter? selected = null;
        string selectedReason = string.Empty;
        int selectedConfidence = 0;
        foreach (ILegacySpreadsheetAdapter adapter in Adapters) {
            cancellationToken.ThrowIfCancellationRequested();
            int confidence = adapter.Probe(data, options.SourceName, options.Limits, cancellationToken, out string reason);
            if (confidence > selectedConfidence) {
                selected = adapter;
                selectedReason = reason;
                selectedConfidence = confidence;
            }
        }
        if (selected == null || selectedConfidence < 50) {
            throw new InvalidDataException("The source does not match a supported bounded legacy-spreadsheet profile. Supply FormatHint only when the family is known.");
        }
        return (selected, new LegacySpreadsheetDetection(selected.Format, selected.GetProfileId(data, options.Limits, cancellationToken), selectedConfidence, selectedReason));
    }

    private static LegacySpreadsheetImportOptions Prepare(LegacySpreadsheetImportOptions? source, string? fallbackName) {
        var options = new LegacySpreadsheetImportOptions {
            Limits = (source?.Limits ?? new OfficeLegacyImportLimits()).Clone(),
            FormatHint = source?.FormatHint,
            SourceName = string.IsNullOrWhiteSpace(source?.SourceName) ? fallbackName : source!.SourceName,
            RequireStructured = source?.RequireStructured ?? false
        };
        options.Limits.Validate();
        return options;
    }
}
