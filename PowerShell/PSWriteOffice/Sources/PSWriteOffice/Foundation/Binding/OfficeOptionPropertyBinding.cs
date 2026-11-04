using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;

namespace PSWriteOffice.Foundation;

/// <summary>Applies typed options or flat overrides only to slots admitted by the operation owner.</summary>
internal static class OfficeOptionPropertyBinding {
    internal static void Apply(EngineIntrinsics engine, object? value,
        IReadOnlyList<(object Owner, PropertyInfo Property)> slots) {
        if (value is IDictionary properties) {
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var slot in slots) {
                var supported = new HashSet<string>(slot.Property.PropertyType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(property => property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0)
                    .Select(property => property.Name), StringComparer.OrdinalIgnoreCase);
                var matching = new Hashtable(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry entry in properties) {
                    if (entry.Key is string key && supported.Contains(key)) { matching[key] = entry.Value; used.Add(key); }
                }
                if (matching.Count == 0) continue;
                // A single flat table supplies complete settings for each matching adapter, rather than mutating caller objects.
                var binder = new OfficeOptionsTransformationAttribute(slot.Property.PropertyType);
                object bound = binder.ApplyTo(engine, slot.Property.GetValue(slot.Owner) ?? Activator.CreateInstance(slot.Property.PropertyType)!, matching);
                slot.Property.SetValue(slot.Owner, bound);
            }
            foreach (DictionaryEntry entry in properties)
                if (entry.Key is not string key || !used.Contains(key))
                    throw new PSArgumentException("No selected conversion route accepts option '" + entry.Key + "'.");
            if (properties.Count == 0) throw new PSArgumentException("Conversion options cannot be an empty hashtable.");
            return;
        }
        var matches = slots.Where(slot => value != null && slot.Property.PropertyType.IsInstanceOfType(value)).ToArray();
        if (matches.Length != 1) throw new PSArgumentException("No single selected conversion option slot accepts '" + (value?.GetType().FullName ?? "<null>") + "'.");
        matches[0].Property.SetValue(matches[0].Owner, value);
    }
}
