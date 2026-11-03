// Vendored from Polyfill (https://github.com/SimonCropp/Polyfill), MIT License, Copyright (c) Simon Cropp.
#if !NET8_0_OR_GREATER
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace Polyfills {
    internal static partial class Polyfill {
        extension(ArgumentException) {
            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentexception.throwifnullorempty#system-argumentexception-throwifnullorempty(system-string-system-string)
            public static void ThrowIfNullOrEmpty([NotNull] string? argument, [CallerArgumentExpression(nameof(argument))] string? paramName = null) {
                if (argument is null) {
                    throw new ArgumentNullException(paramName);
                }

                if (argument.Length == 0) {
                    throw new ArgumentException("The value cannot be an empty string.", paramName);
                }
            }

            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentexception.throwifnullorwhitespace#system-argumentexception-throwifnullorwhitespace(system-string-system-string)
            public static void ThrowIfNullOrWhiteSpace([NotNull] string? argument, [CallerArgumentExpression(nameof(argument))] string? paramName = null) {
                if (argument is null) {
                    throw new ArgumentNullException(paramName);
                }

                if (string.IsNullOrWhiteSpace(argument)) {
                    throw new ArgumentException("The value cannot be an empty string or composed entirely of whitespace.", paramName);
                }
            }
        }
    }
}
#endif
