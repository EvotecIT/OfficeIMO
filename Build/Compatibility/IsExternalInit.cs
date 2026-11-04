#if NETSTANDARD2_0 || NETFRAMEWORK
namespace System.Runtime.CompilerServices {
    internal static class IsExternalInit {
    }
}
#endif

#if NETSTANDARD2_0 || NETFRAMEWORK
namespace System.Runtime.CompilerServices {
    [System.AttributeUsage(System.AttributeTargets.Class | System.AttributeTargets.Struct | System.AttributeTargets.Field | System.AttributeTargets.Property, Inherited = false)]
    internal sealed class RequiredMemberAttribute : System.Attribute { }

    [System.AttributeUsage(System.AttributeTargets.All, AllowMultiple = true, Inherited = false)]
    internal sealed class CompilerFeatureRequiredAttribute : System.Attribute {
        public CompilerFeatureRequiredAttribute(string featureName) { FeatureName = featureName; }
        public string FeatureName { get; }
        public bool IsOptional { get; init; }
    }
}
#endif

#if NETSTANDARD2_0 || NETFRAMEWORK
namespace System.Diagnostics.CodeAnalysis {
    [System.AttributeUsage(System.AttributeTargets.Constructor, Inherited = false)]
    internal sealed class SetsRequiredMembersAttribute : System.Attribute { }
}
#endif
