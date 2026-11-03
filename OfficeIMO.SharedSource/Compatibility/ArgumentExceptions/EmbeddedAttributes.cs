// Vendored from Polyfill (https://github.com/SimonCropp/Polyfill), MIT License, Copyright (c) Simon Cropp.
// Every shim is marked [Embedded] so it stays invisible to InternalsVisibleTo friends that compile their own copy.
#if !NET8_0_OR_GREATER
namespace Microsoft.CodeAnalysis {
    [Embedded]
    [System.AttributeUsage(System.AttributeTargets.All, Inherited = false)]
    internal sealed class EmbeddedAttribute : System.Attribute {
    }
}

#if !NET6_0_OR_GREATER
namespace System.Runtime.CompilerServices {
    [Microsoft.CodeAnalysis.Embedded]
    [AttributeUsage(AttributeTargets.Parameter)]
    internal sealed class CallerArgumentExpressionAttribute : Attribute {
        public CallerArgumentExpressionAttribute(string parameterName) => ParameterName = parameterName;

        public string ParameterName { get; }
    }
}
#endif

#if !NETCOREAPP3_0_OR_GREATER && !NETSTANDARD2_1_OR_GREATER
namespace System.Diagnostics.CodeAnalysis {
    [Microsoft.CodeAnalysis.Embedded]
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Parameter | AttributeTargets.Property | AttributeTargets.ReturnValue)]
    internal sealed class NotNullAttribute : Attribute {
    }
}
#endif
#endif
