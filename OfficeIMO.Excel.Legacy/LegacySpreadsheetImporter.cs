using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

    /// <summary>Detects legacy spreadsheet bytes.</summary>
    public static LegacySpreadsheetDetection Detect(byte[] data, LegacySpreadsheetImportOptions? options = null, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        LegacySpreadsheetImportOptions effective = Prepare(options, options?.SourceName);
        if (data.Length > effective.Limits.MaxInputBytes) throw new InvalidDataException("Legacy spreadsheet input exceeds the configured byte limit.");
        cancellationToken.ThrowIfCancellationRequested();
        return SelectAdapter(data, effective).Detection;
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
        (ILegacySpreadsheetAdapter adapter, LegacySpreadsheetDetection detection) = SelectAdapter(data, effective);
        cancellationToken.ThrowIfCancellationRequested();
        LegacySpreadsheetModel model = adapter.Parse(data, effective.Limits, cancellationToken);
        if (effective.RequireStructured && model.Quality != OfficeLegacyImportQuality.Structured) {
            throw new InvalidDataException($"The {detection.ProfileId} adapter produced salvage quality while structured import was required.");
        }

        ExcelDocument document = Project(model, cancellationToken);
        var report = new OfficeLegacyImportReport(detection.ProfileId, model.Quality, model.Findings, model.InertContent, model.RecoveredCellCount);
        return new LegacySpreadsheetImportResult(document, detection, report,
            new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(model.Metadata, StringComparer.OrdinalIgnoreCase)),
            Array.AsReadOnly(model.Charts.ToArray()));
    }

    private static ExcelDocument Project(LegacySpreadsheetModel model, CancellationToken cancellationToken) {
        ExcelDocument document = ExcelDocument.Create();
        try {
            foreach (LegacySpreadsheetSheet sourceSheet in model.Sheets) {
                cancellationToken.ThrowIfCancellationRequested();
                ExcelSheet target = document.AddWorksheet(sourceSheet.Name);
                foreach (LegacySpreadsheetCell cell in sourceSheet.Cells) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.IsNullOrWhiteSpace(cell.Formula)) target.CellFormula(cell.Row, cell.Column, cell.Formula!);
                    else target.CellValue(cell.Row, cell.Column, cell.Value);
                    if (cell.Alignment.HasValue) target.CellAlign(cell.Row, cell.Column, cell.Alignment.Value);
                    if (!string.IsNullOrWhiteSpace(cell.Comment)) target.SetComment(cell.Row, cell.Column, cell.Comment!, "Legacy source");
                }
            }
            if (model.Sheets.Count == 0) document.AddWorksheet("Sheet1");
            return document;
        } catch {
            document.Dispose();
            throw;
        }
    }

    private static (ILegacySpreadsheetAdapter Adapter, LegacySpreadsheetDetection Detection) SelectAdapter(byte[] data, LegacySpreadsheetImportOptions options) {
        if (options.FormatHint.HasValue) {
            ILegacySpreadsheetAdapter hinted = Adapters.Single(adapter => adapter.Format == options.FormatHint.Value);
            int confidence = hinted.Probe(data, options.SourceName, out string evidence);
            return (hinted, new LegacySpreadsheetDetection(hinted.Format, hinted.ProfileId, Math.Max(1, confidence),
                confidence == 0 ? "Explicit caller format hint." : evidence + " Explicit caller format hint confirmed the family."));
        }
        ILegacySpreadsheetAdapter? selected = null;
        string selectedReason = string.Empty;
        int selectedConfidence = 0;
        foreach (ILegacySpreadsheetAdapter adapter in Adapters) {
            int confidence = adapter.Probe(data, options.SourceName, out string reason);
            if (confidence > selectedConfidence) {
                selected = adapter;
                selectedReason = reason;
                selectedConfidence = confidence;
            }
        }
        if (selected == null || selectedConfidence < 50) {
            throw new InvalidDataException("The source does not match a supported bounded legacy-spreadsheet profile. Supply FormatHint only when the family is known.");
        }
        return (selected, new LegacySpreadsheetDetection(selected.Format, selected.ProfileId, selectedConfidence, selectedReason));
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
