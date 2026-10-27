using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Management.Automation;
using System.Reflection;

namespace PSWriteOffice.Foundation;

/// <summary>Maps a hashtable to an existing typed options contract, rejecting unknown or read-only properties.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficeOptionsTransformationAttribute : ArgumentTransformationAttribute {
    /// <summary>Creates a binder for a concrete options type with a public parameterless constructor.</summary>
    public OfficeOptionsTransformationAttribute(Type optionsType) {
        OptionsType = optionsType ?? throw new ArgumentNullException(nameof(optionsType));
        if (optionsType.IsAbstract || optionsType.IsInterface || optionsType.GetConstructor(Type.EmptyTypes) == null)
            throw new ArgumentException("Options must have a public parameterless constructor.", nameof(optionsType));
    }
    /// <summary>Gets the command's options type.</summary>
    public Type OptionsType { get; }
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value != null && OptionsType.IsInstanceOfType(value)) return value;
        if (value is not IDictionary dictionary) throw OfficeArgumentValue.Invalid("Expected a typed options object or hashtable.");
        var properties = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in OptionsType.GetProperties(BindingFlags.Public | BindingFlags.Instance)) {
            if (property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0) properties[property.Name] = property;
        }
        object result = Activator.CreateInstance(OptionsType)!;
        foreach (DictionaryEntry entry in dictionary) {
            if (entry.Key is not string key || !properties.TryGetValue(key, out var property))
                throw OfficeArgumentValue.Invalid($"Unknown or read-only {OptionsType.Name} option '{entry.Key}'.");
            try {
                var input = OfficeArgumentValue.Unwrap(entry.Value);
                object? converted = input;
                // Reuse the same binding contracts for nested common values.
                if (property.PropertyType == typeof(OfficeIMO.OfficeLength)) converted = new OfficeLengthTransformationAttribute().Transform(engineIntrinsics, input!);
                else if (property.PropertyType == typeof(OfficeIMO.OfficeColor)) converted = new OfficeColorTransformationAttribute().Transform(engineIntrinsics, input!);
                else converted = LanguagePrimitives.ConvertTo(input, property.PropertyType, CultureInfo.InvariantCulture);
                property.SetValue(result, converted);
            } catch (Exception error) when (error is PSInvalidCastException or ArgumentException or TargetInvocationException) {
                throw OfficeArgumentValue.Invalid($"Invalid value for {OptionsType.Name}.{property.Name}: {error.InnerException?.Message ?? error.Message}");
            }
        }
        return result;
    }
}
