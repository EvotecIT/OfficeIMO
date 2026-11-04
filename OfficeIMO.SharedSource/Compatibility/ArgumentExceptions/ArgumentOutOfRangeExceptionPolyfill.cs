// Vendored from Polyfill (https://github.com/SimonCropp/Polyfill), MIT License, Copyright (c) Simon Cropp.
// Only the generic overloads are vendored; the INumberBase<T> variants are unnecessary because the file compiles only below net8.0.
#if !NET8_0_OR_GREATER
using System;
using System.Runtime.CompilerServices;

namespace Polyfills {
    internal static partial class Polyfill {
        extension(ArgumentOutOfRangeException) {
            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception.throwifnegative
            public static void ThrowIfNegative<T>(T value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
                where T : struct, IComparable<T> {
                if (value.CompareTo(default) < 0) {
                    throw new ArgumentOutOfRangeException(paramName, value, "Value must be non-negative.");
                }
            }

            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception.throwifnegativeorzero
            public static void ThrowIfNegativeOrZero<T>(T value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
                where T : struct, IComparable<T> {
                if (value.CompareTo(default) <= 0) {
                    throw new ArgumentOutOfRangeException(paramName, value, $"{paramName} ('{value}') must be a non-negative and non-zero value.");
                }
            }

            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception.throwifgreaterthan
            public static void ThrowIfGreaterThan<T>(T value, T other, [CallerArgumentExpression(nameof(value))] string? paramName = null)
                where T : IComparable<T> {
                if (value.CompareTo(other) > 0) {
                    throw new ArgumentOutOfRangeException(paramName, value, $"Value must be less than or equal to {other}.");
                }
            }

            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception.throwifgreaterthanorequal
            public static void ThrowIfGreaterThanOrEqual<T>(T value, T other, [CallerArgumentExpression(nameof(value))] string? paramName = null)
                where T : IComparable<T> {
                if (value.CompareTo(other) >= 0) {
                    throw new ArgumentOutOfRangeException(paramName, value, $"Value must be less than {other}.");
                }
            }

            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception.throwiflessthan
            public static void ThrowIfLessThan<T>(T value, T other, [CallerArgumentExpression(nameof(value))] string? paramName = null)
                where T : IComparable<T> {
                if (value.CompareTo(other) < 0) {
                    throw new ArgumentOutOfRangeException(paramName, value, $"Value must be greater than or equal to {other}.");
                }
            }

            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception.throwiflessthanorequal
            public static void ThrowIfLessThanOrEqual<T>(T value, T other, [CallerArgumentExpression(nameof(value))] string? paramName = null)
                where T : IComparable<T> {
                if (value.CompareTo(other) <= 0) {
                    throw new ArgumentOutOfRangeException(paramName, value, $"Value must be greater than {other}.");
                }
            }
        }
    }
}
#endif
