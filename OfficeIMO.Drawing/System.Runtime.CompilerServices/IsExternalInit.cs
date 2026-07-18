#if NETSTANDARD2_0 || NETFRAMEWORK
namespace System.Runtime.CompilerServices;

using System.ComponentModel;

/// <summary>
/// Reserved for compiler metadata that supports init-only properties on older targets.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class IsExternalInit { }
#endif
