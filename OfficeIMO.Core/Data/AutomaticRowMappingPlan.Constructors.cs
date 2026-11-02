#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace OfficeIMO.Data;

internal sealed partial class AutomaticRowMappingPlan<T> {
    // An explicit public parameterless constructor retains its existing path.
    // Constructor-bound classes and values use the resolved property/alias plan.
    private static ConstructorBinding? CreateConstructorBinding(MappingBinding[] bindings) {
        if (typeof(T).GetConstructor(Type.EmptyTypes) is not null) return null;

        ConstructorBinding? best = null;
        int bestScore = -1;
        bool ambiguous = false;
        ConstructorInfo[] constructors = typeof(T).GetConstructors();
        foreach (ConstructorInfo constructor in constructors) {
            ParameterInfo[] parameters = constructor.GetParameters();
            var arguments = new MappingBinding?[parameters.Length];
            int score = 0;
            bool applicable = true;
            for (int index = 0; index < parameters.Length; index++) {
                ParameterInfo parameter = parameters[index];
                MappingBinding[] matches = bindings.Where(binding =>
                    string.Equals(binding.Property.Name, parameter.Name, StringComparison.Ordinal)).ToArray();
                if (matches.Length == 0) {
                    matches = bindings.Where(binding =>
                        string.Equals(binding.Property.Name, parameter.Name, StringComparison.OrdinalIgnoreCase)).ToArray();
                }
                if (matches.Length == 1 && matches[0].Property.ValueType == parameter.ParameterType) {
                    arguments[index] = matches[0];
                    score++;
                } else if (matches.Length > 0 || !parameter.HasDefaultValue) {
                    applicable = false;
                    break;
                }
            }
            if (!applicable || score < bestScore) continue;
            if (score == bestScore) {
                ambiguous = true;
                continue;
            }
            best = new ConstructorBinding(constructor, parameters, arguments);
            bestScore = score;
            ambiguous = false;
        }
        if (best is null) {
            if (typeof(T).IsValueType && constructors.Length == 0
                && bindings.All(binding => binding.Property.CanAssign)) return null;
            throw new DataMappingException(
                $"No public constructor on {typeof(T).Name} matches the mapped columns. Supply a row factory for this model.");
        }
        if (ambiguous) {
            throw new DataMappingException(
                $"Multiple public constructors on {typeof(T).Name} match the mapped columns. Supply a row factory to choose one.");
        }
        return best;
    }

    private static bool CanBindConstructorProperty(PropertyInfo property) {
        if (property.GetMethod?.IsPublic != true || typeof(T).GetConstructor(Type.EmptyTypes) is not null) return false;
        return typeof(T).GetConstructors().Any(constructor => constructor.GetParameters().Any(parameter =>
            string.Equals(parameter.Name, property.Name, StringComparison.OrdinalIgnoreCase)
            && parameter.ParameterType == property.PropertyType));
    }

    private sealed class ConstructorBinding {
        private readonly ConstructorInfo _constructor;
        private readonly ParameterInfo[] _parameters;
        private readonly MappingBinding?[] _arguments;

        internal ConstructorBinding(
            ConstructorInfo constructor,
            ParameterInfo[] parameters,
            MappingBinding?[] arguments) {
            _constructor = constructor;
            _parameters = parameters;
            _arguments = arguments;
        }

        internal bool Contains(AutomaticPropertySetter property) =>
            _arguments.Any(argument => argument.HasValue && ReferenceEquals(argument.Value.Property, property));

        internal T Create(
            Func<MappingBinding, object?> getValue,
            CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats,
            Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            DataMappingErrorValuePolicy errorValuePolicy) {
            var values = new object?[_parameters.Length];
            for (int index = 0; index < values.Length; index++) {
                ParameterInfo parameter = _parameters[index];
                if (_arguments[index] is not MappingBinding binding) {
                    values[index] = parameter.DefaultValue;
                    continue;
                }
                object? raw = getValue(binding);
                if (!DataValueConverter.TryConvert(
                        ReferenceEquals(raw, DBNull.Value) ? null : raw,
                        parameter.ParameterType, culture, dateTimeFormats, typeConverter, errorValuePolicy,
                        out object? converted, out string? error)) {
                    throw new DataMappingException(
                        $"Column value cannot be assigned to {typeof(T).Name}.{binding.Property.Name}: {error}");
                }
                values[index] = converted;
            }
            try {
                return (T)_constructor.Invoke(values);
            } catch (TargetInvocationException exception) {
                string detail = errorValuePolicy == DataMappingErrorValuePolicy.Redact
                    ? string.Empty : $": {exception.InnerException?.Message ?? exception.Message}";
                throw new DataMappingException($"Constructor for {typeof(T).Name} rejected the mapped row{detail}.");
            }
        }
    }
}
