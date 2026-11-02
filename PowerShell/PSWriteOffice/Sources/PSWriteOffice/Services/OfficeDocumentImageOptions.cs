using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

/// <summary>Snapshots only the native image and load contracts used by admitted source formats.</summary>
internal sealed class OfficeDocumentImageOptions {
    public WordImageExportOptions Word { get; set; } = new();
    public ExcelImageExportOptions Excel { get; set; } = new();
    public PowerPointImageExportOptions PowerPoint { get; set; } = new();
    public PdfImageExportOptions Pdf { get; set; } = new();
    public VisioImageExportOptions Visio { get; set; } = new();
    public HtmlRenderOptions Html { get; set; } = new();
    public WordLoadOptions WordLoad { get; set; } = new();
    public ExcelLoadOptions ExcelLoad { get; set; } = new();
    public PowerPointLoadOptions PowerPointLoad { get; set; } = new();
    public PdfLoadOptions PdfLoad { get; set; } = new();
    public VisioLoadOptions VisioLoad { get; set; } = new();
    public HtmlConversionDocumentOptions HtmlLoad { get; set; } = new();

    internal static OfficeDocumentImageOptions Capture(EngineIntrinsics engine, object[]? values, IReadOnlyCollection<OfficeDocumentFormat> formats) {
        var result = new OfficeDocumentImageOptions();
        foreach (object? item in values ?? Array.Empty<object>()) {
            object? value = OfficeArgumentValue.Unwrap(item);
            if (value is IDictionary dictionary && dictionary.Count > 0 && dictionary.Keys.Cast<object>()
                .All(key => key is string text && OfficeFormatSelection.TryParse(text, out _))) {
                foreach (DictionaryEntry entry in dictionary) {
                    OfficeDocumentFormat format = (OfficeFormatSelection.Parse((string)entry.Key)!.Family ?? OfficeConverter.FormatForExtension(OfficeFormatSelection.Parse((string)entry.Key)!.Extension));
                    if (!formats.Contains(format)) throw new PSArgumentException("No selected image source uses format '" + entry.Key + "'.");
                    Apply(engine, result, OfficeArgumentValue.Unwrap(entry.Value), new[] { format });
                }
            } else Apply(engine, result, value, formats);
            result = result.Clone();
        }
        return result;
    }

    private static void Apply(EngineIntrinsics engine, OfficeDocumentImageOptions result, object? value, IReadOnlyCollection<OfficeDocumentFormat> formats) {
        var names = new HashSet<string>(formats.Select(format => format.ToString()), StringComparer.Ordinal);
        var slots = typeof(OfficeDocumentImageOptions).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(property => names.Contains(property.Name.EndsWith("Load", StringComparison.Ordinal) ? property.Name.Substring(0, property.Name.Length - 4) : property.Name))
            .Select(property => ((object)result, property)).ToArray();
        OfficeOptionPropertyBinding.Apply(engine, value, slots);
    }

    internal OfficeImageExportOptions Image(OfficeDocumentFormat format) => format switch {
        OfficeDocumentFormat.Word => Word,
        OfficeDocumentFormat.Excel => Excel,
        OfficeDocumentFormat.PowerPoint => PowerPoint,
        OfficeDocumentFormat.Pdf => Pdf,
        OfficeDocumentFormat.Visio => Visio,
        OfficeDocumentFormat.Html => Html,
        _ => throw new PSArgumentException("This source format has no native document image exporter.")
    };

    internal OfficeDocumentImageOptions Clone() => new() {
        Word = Word.Clone(), Excel = Excel.Clone(), PowerPoint = PowerPoint.Clone(), Pdf = Pdf.Clone(),
        Visio = Visio.Clone(), Html = Html.Clone(), WordLoad = WordLoad.Snapshot(), ExcelLoad = ExcelLoad.Snapshot(),
        PowerPointLoad = PowerPointLoad.Snapshot(), PdfLoad = PdfLoad.Clone(), HtmlLoad = HtmlLoad.Clone(),
        VisioLoad = new VisioLoadOptions {
            MaxInputBytes = VisioLoad.MaxInputBytes, MaxLegacyXmlCharacters = VisioLoad.MaxLegacyXmlCharacters,
            MaxPackageXmlCharacters = VisioLoad.MaxPackageXmlCharacters, MaxLegacyXmlDepth = VisioLoad.MaxLegacyXmlDepth,
            MaxLegacyXmlElements = VisioLoad.MaxLegacyXmlElements, PackageSecurity = VisioLoad.PackageSecurity?.Snapshot()
        }
    };
}
