using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO;
using OfficeIMO.Workflows;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

/// <summary>Routes PowerShell option values to existing adapter-owned option properties.</summary>
internal static class OfficeConversionOptionBinding {
    internal static OfficeConversionOptions Capture(EngineIntrinsics engine, object[]? values,
        IReadOnlyList<OfficeConversionCapability> routes) {
        var result = new OfficeConversionOptions();
        foreach (object? item in values ?? Array.Empty<object>()) {
            object? value = Unwrap(item);
            if (value is OfficeConversionOptions options) {
                if (values!.Length != 1) throw new PSArgumentException("A complete OfficeConversionOptions object cannot be combined with other option values.");
                return options.Clone();
            }
            Apply(engine, result, value, routes);
            result = result.Clone();
        }
        return result.Clone();
    }

    private static void Apply(EngineIntrinsics engine, OfficeConversionOptions destination, object? value,
        IReadOnlyList<OfficeConversionCapability> routes) {
        if (value is IDictionary dictionary && IsFormatDictionary(dictionary)) {
            foreach (DictionaryEntry entry in dictionary) {
                var format = OfficeFormatSelection.Parse((string)entry.Key)!;
                var selected = routes.Where(route => route.SourceExtensions.Any(format.Matches)).ToArray();
                if (selected.Length == 0) throw new PSArgumentException("No selected conversion route uses source format '" + entry.Key + "'.");
                Apply(engine, destination, Unwrap(entry.Value), selected);
            }
            return;
        }
        OfficeOptionPropertyBinding.Apply(engine, value, GetSlots(destination, routes).ToArray());
    }

    private static bool IsFormatDictionary(IDictionary dictionary) => dictionary.Count > 0 &&
        dictionary.Keys.Cast<object>().All(key => key is string text && OfficeFormatSelection.TryParse(text, out _));

    private static IEnumerable<(object Owner, PropertyInfo Property)> GetSlots(OfficeConversionOptions options,
        IReadOnlyList<OfficeConversionCapability> routes) {
        // ForRoute is the engine's canonical applicability policy. Sentinel objects ask that policy which typed slots survive.
        var candidates = typeof(OfficeConversionOptions).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(property => property.SetMethod?.IsPublic == true && property.PropertyType.IsClass &&
                property.PropertyType != typeof(string) && property.PropertyType.GetConstructor(Type.EmptyTypes) != null).ToArray();
        var populated = new OfficeConversionOptions();
        foreach (var property in candidates) property.SetValue(populated, Activator.CreateInstance(property.PropertyType));
        var accepted = new HashSet<string>(StringComparer.Ordinal);
        foreach (var route in routes) {
            var selected = populated.ForRoute(route.Id);
            foreach (var property in candidates) if (property.GetValue(selected) != null) accepted.Add(property.Name);
        }
        foreach (var property in candidates.Where(property => accepted.Contains(property.Name)))
            yield return (options, property);
        var registered = routes.Where(route => !OfficeConverter.IsBuiltinRoute(route.Id)).ToArray();
        if (registered.Length == 0) yield break;
        options.RegisteredSettings ??= new OfficeWorkflowRouteSettings();
        if (options.RegisteredSettings is not OfficeWorkflowRouteSettings settings) yield break;
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var route in registered) {
            switch (route.Id) {
                case "rtf-pdf": names.Add(nameof(settings.Rtf)); break;
                case "odg-pdf": names.Add(nameof(settings.Draw)); break;
                case "visio-pdf": names.Add(nameof(settings.Visio)); break;
                case "xps-pdf": names.Add(nameof(settings.Xps)); break;
                case "djvu-pdf": names.Add(nameof(settings.DjVu)); break;
                case "chm-pdf": names.Add(nameof(settings.ChmHtml)); names.Add(nameof(settings.Chm)); break;
                case "chm-markdown":
                case "chm-epub": names.Add(nameof(settings.Chm)); break;
                case "publisher-pdf": names.Add(nameof(settings.PublisherRead)); names.Add(nameof(settings.PublisherPdf)); break;
            }
        }
        foreach (var property in typeof(OfficeWorkflowRouteSettings).GetProperties().Where(property => names.Contains(property.Name)))
            yield return (settings, property);
    }

    private static object? Unwrap(object? value) {
        while (value is PSObject wrapped && !ReferenceEquals(value, wrapped.BaseObject)) value = wrapped.BaseObject;
        return value;
    }
}
