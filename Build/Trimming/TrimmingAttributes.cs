#if !NET5_0_OR_GREATER
using System;

namespace System.Diagnostics.CodeAnalysis {
    [Flags]
    internal enum DynamicallyAccessedMemberTypes {
        None = 0,
        PublicParameterlessConstructor = 0x0001,
        PublicConstructors = 0x0002 | PublicParameterlessConstructor,
        NonPublicConstructors = 0x0004,
        PublicMethods = 0x0008,
        NonPublicMethods = 0x0010,
        PublicFields = 0x0020,
        NonPublicFields = 0x0040,
        PublicNestedTypes = 0x0080,
        NonPublicNestedTypes = 0x0100,
        PublicProperties = 0x0200,
        NonPublicProperties = 0x0400,
        PublicEvents = 0x0800,
        NonPublicEvents = 0x1000,
        All = ~0
    }

    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct | AttributeTargets.Interface |
                    AttributeTargets.Delegate | AttributeTargets.Method | AttributeTargets.Property |
                    AttributeTargets.Field | AttributeTargets.Event | AttributeTargets.Parameter |
                    AttributeTargets.ReturnValue | AttributeTargets.GenericParameter, Inherited = false)]
    internal sealed class DynamicallyAccessedMembersAttribute : Attribute {
        public DynamicallyAccessedMembersAttribute(DynamicallyAccessedMemberTypes memberTypes) {
            MemberTypes = memberTypes;
        }

        public DynamicallyAccessedMemberTypes MemberTypes { get; }
    }

    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class, Inherited = false)]
    internal sealed class RequiresUnreferencedCodeAttribute : Attribute {
        public RequiresUnreferencedCodeAttribute(string message) {
            Message = message;
        }

        public string Message { get; }
        public string? Url { get; set; }
    }

    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class, Inherited = false)]
    internal sealed class RequiresDynamicCodeAttribute : Attribute {
        public RequiresDynamicCodeAttribute(string message) {
            Message = message;
        }

        public string Message { get; }
        public string? Url { get; set; }
    }

    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class, Inherited = false)]
    internal sealed class RequiresAssemblyFilesAttribute : Attribute {
        public RequiresAssemblyFilesAttribute(string message) {
            Message = message;
        }

        public string Message { get; }
        public string? Url { get; set; }
    }

    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
    internal sealed class DynamicDependencyAttribute : Attribute {
        public DynamicDependencyAttribute(string memberSignature) {
            MemberSignature = memberSignature;
        }

        public DynamicDependencyAttribute(string memberSignature, Type type) {
            MemberSignature = memberSignature;
            Type = type;
        }

        public string MemberSignature { get; }
        public Type? Type { get; }
    }
}
#endif
