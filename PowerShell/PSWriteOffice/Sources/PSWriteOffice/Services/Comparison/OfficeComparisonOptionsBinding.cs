using System.Management.Automation;
using System.Reflection;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Comparison;

/// <summary>Captures comparison inputs while leaving caller-owned native settings unchanged.</summary>
internal static class OfficeComparisonOptionsBinding {
    internal static T Capture<T>(EngineIntrinsics engine, object[]? inputs) where T : class, new() {
        var result = new T();
        foreach (var input in inputs ?? Array.Empty<object>()) {
            var value = OfficeArgumentValue.Unwrap(input);
            if (value is T typed) result = Copy(typed);
            else if (value is System.Collections.IDictionary properties)
                result = (T)new OfficeOptionsTransformationAttribute(typeof(T)).ApplyTo(engine, result, properties);
            else throw new PSArgumentException("Options must be " + typeof(T).Name + " or native property overrides.");
        }
        return result;
    }
    internal static OfficePdfVisualComparisonOptions CapturePdf(EngineIntrinsics engine, object[]? inputs) {
        var result = new OfficePdfVisualComparisonOptions();
        foreach (var input in inputs ?? Array.Empty<object>()) {
            var value = OfficeArgumentValue.Unwrap(input);
            if (value is OfficePdfVisualComparisonOptions typed) result = new OfficePdfVisualComparisonOptions {
                Comparison = Copy(typed.Comparison), ReferenceLoadOptions = typed.ReferenceLoadOptions?.Clone(), DifferenceLoadOptions = typed.DifferenceLoadOptions?.Clone()
            };
            else if (value is PdfVisualComparisonOptions native) result.Comparison = Copy(native);
            else if (value is PdfLoadOptions load) { result.ReferenceLoadOptions = load.Clone(); result.DifferenceLoadOptions = load.Clone(); }
            else if (value is System.Collections.IDictionary properties)
                result.Comparison = (PdfVisualComparisonOptions)new OfficeOptionsTransformationAttribute(typeof(PdfVisualComparisonOptions)).ApplyTo(engine, result.Comparison, properties);
            else throw new PSArgumentException("Visual Options must be native PDF comparison/read settings or their property overrides.");
        }
        return result;
    }
    // Comparison options are input-only values; setters receive an independent container before any override.
    private static T Copy<T>(T source) where T : class, new() {
        var result = new T();
        foreach (var property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance))
            if (property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0) property.SetValue(result, property.GetValue(source));
        if (source is PdfVisualComparisonOptions pdf && result is PdfVisualComparisonOptions copy)
            foreach (var region in pdf.IgnoredRegions) copy.IgnoredRegions.Add(region);
        return result;
    }
}
