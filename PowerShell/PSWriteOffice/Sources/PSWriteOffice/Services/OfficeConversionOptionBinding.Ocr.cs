using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

internal static partial class OfficeConversionOptionBinding {
    /// <summary>Binds the existing native PDF source and OCR projection options without inventing route registrations.</summary>
    internal static OfficeConversionOptions CaptureOcr(EngineIntrinsics engine, object[]? values, string target) {
        var result = new OfficeConversionOptions();
        foreach (object? item in values ?? Array.Empty<object>()) {
            object? value = Unwrap(item);
            if (value is OfficeConversionOptions options) {
                if (values!.Length != 1) throw new PSArgumentException("A complete OfficeConversionOptions object cannot be combined with other option values.");
                return options.Clone();
            }
            ApplyOcr(engine, result, value, target);
            result = result.Clone();
        }
        return result;
    }

    private static void ApplyOcr(EngineIntrinsics engine, OfficeConversionOptions options, object? value, string target) {
        if (value is Models.Pdf.PdfReadOptions pdf) {
            var captured = pdf.Snapshot();
            if (captured.OcrMergeOptions != null) throw new PSArgumentException("Recognition settings belong in OcrOptions.");
            value = captured.LoadOptions;
        }
        if (value is IDictionary dictionary && IsFormatDictionary(dictionary)) {
            foreach (DictionaryEntry entry in dictionary) {
                if (OfficeFormatSelection.Parse((string)entry.Key)!.Family != OfficeDocumentFormat.Pdf)
                    throw new PSArgumentException("OCR projection options use the normalized Pdf source family.");
                ApplyOcr(engine, options, Unwrap(entry.Value), target);
            }
            return;
        }
        string? projection = target.ToLowerInvariant() switch {
            ".docx" => nameof(OfficeConversionOptions.PdfWord),
            ".xlsx" => nameof(OfficeConversionOptions.PdfExcel),
            ".html" => nameof(OfficeConversionOptions.PdfHtml),
            ".csv" => nameof(OfficeConversionOptions.CsvSave),
            _ => null
        };
        string[] names = projection == null
            ? new[] { nameof(OfficeConversionOptions.PdfLoad), nameof(OfficeConversionOptions.Reader) }
            : new[] { nameof(OfficeConversionOptions.PdfLoad), nameof(OfficeConversionOptions.Reader), projection };
        OfficeOptionPropertyBinding.Apply(engine, value, names.Select(name =>
            ((object)options, typeof(OfficeConversionOptions).GetProperty(name, BindingFlags.Public | BindingFlags.Instance)!)).ToArray());
    }
}
