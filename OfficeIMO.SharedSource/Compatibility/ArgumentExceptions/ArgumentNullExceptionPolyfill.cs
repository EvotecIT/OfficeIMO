// Vendored from Polyfill (https://github.com/SimonCropp/Polyfill), MIT License, Copyright (c) Simon Cropp.
#if !NET8_0_OR_GREATER
global using Polyfills;

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace Polyfills {
    [Microsoft.CodeAnalysis.Embedded]
    internal static partial class Polyfill {
        extension(ArgumentNullException) {
#if !NET6_0_OR_GREATER
            //Link: https://learn.microsoft.com/en-us/dotnet/api/system.argumentnullexception.throwifnull#system-argumentnullexception-throwifnull(system-object-system-string)
            public static void ThrowIfNull([NotNull] object? argument, [CallerArgumentExpression(nameof(argument))] string? paramName = null) {
                if (argument is null) {
                    throw new ArgumentNullException(paramName);
                }
            }
#endif
        }
    }
}
#endif
