using System;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace PSWriteOffice.Services.Reader;

/// <summary>Binding-only option routing; native Reader handlers own validation and defensive snapshots.</summary>
internal sealed class OfficeReadConfiguration {
    internal ReaderOptions Common { get; private set; } = new();
    internal ReaderAllOptions Formats { get; private set; } = new();
    internal bool HasCommonOptions { get; private set; }
    internal bool HasFormatOptions { get; private set; }
    internal OfficeIMO.Pdf.Ocr.PdfOcrMergeOptions? PdfOcrMergeOptions { get; private set; }

    internal static OfficeReadConfiguration Bind(object[]? inputs, ReaderOptions? defaults = null) {
        var result = new OfficeReadConfiguration();
        if (defaults != null) result.Common = Copy(defaults);
        foreach (object input in inputs ?? Array.Empty<object>()) {
            object value = input is PSObject wrapped ? wrapped.BaseObject : input;
            if (value is ReaderOptions commonOptions) { result.Common = Copy(commonOptions); result.HasCommonOptions = true; continue; }
            if (value is ReaderAllOptions all) { result.Formats = Copy(all); result.HasFormatOptions = true; continue; }
            PropertyInfo? property = typeof(ReaderAllOptions).GetProperties().SingleOrDefault(item => item.PropertyType.IsClass && item.PropertyType == value?.GetType());
            if (property != null) { property.SetValue(result.Formats, value); result.HasFormatOptions = true; continue; }
            switch (value) {
                case Models.Pdf.PdfReadOptions composite:
                    var captured = composite.Snapshot();
                    result.Formats.Pdf = Copy(result.Formats.Pdf);
                    result.Formats.Pdf.LoadOptions = captured.LoadOptions;
                    result.Formats.Pdf.ReadOptions = captured.LoadOptions;
                    result.PdfOcrMergeOptions = captured.OcrMergeOptions;
                    break;
                case OfficeIMO.Pdf.PdfLoadOptions pdf:
                    // The canonical PDF policy controls initial admission as well as semantic reconstruction.
                    var policy = pdf.Clone();
                    result.Formats.Pdf = Copy(result.Formats.Pdf);
                    result.Formats.Pdf.LoadOptions = policy;
                    result.Formats.Pdf.ReadOptions = policy;
                    break;
                case OfficeIMO.Markdown.MarkdownLoadOptions markdown:
                    result.Formats.Markdown = Copy(result.Formats.Markdown); result.Formats.Markdown.ParserOptions = markdown; break;
                case OfficeIMO.Excel.ExcelReaderOptions excel:
                    result.Formats.Excel = Copy(result.Formats.Excel); result.Formats.Excel.ReadOptions = excel; break;
                default: throw new PSArgumentException("Unsupported read options type: " + (value?.GetType().FullName ?? "null") + ". Use ReaderOptions, ReaderAllOptions or a registered format's Reader options.");
            }
            result.HasFormatOptions = true;
        }
        return result;
    }

    internal static ReaderFolderOptions CreateFolder(ReaderOptions common, bool? recurse, string[]? extensions, bool extensionsSpecified) {
        var folder = Copy(common.FolderOptions);
        if (common.FolderOptions == null) folder.Recurse = false;
        if (recurse.HasValue) folder.Recurse = recurse.Value;
        if (extensionsSpecified) folder.Extensions = extensions;
        folder.MaxFiles = int.MaxValue;
        // The batch owns total source admission. Enumeration never parses a directory package to estimate its length.
        folder.MaxTotalBytes = null;
        return folder;
    }

    // Copy only the binding object before applying flat overrides. Each registered owner snapshots nested policy itself.
    internal static T Copy<T>(T? value) where T : class, new() {
        var copy = new T();
        if (value != null) foreach (var property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance))
            if (property.CanRead && property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0)
                property.SetValue(copy, property.GetValue(value));
        return copy;
    }
}
