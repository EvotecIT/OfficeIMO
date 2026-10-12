using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Serialization;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Native;

/// <summary>Version, status codes and memory rules of the C interface declared in <c>include/officeimo_studio.h</c>.</summary>
internal static class StudioNativeAbi {
    /// <summary>Raised whenever an entry point's signature or a JSON payload changes incompatibly.</summary>
    internal const int Version = 1;

    internal const int Ok = 0;
    internal const int Failed = 1;
    internal const int Cancelled = 2;
    internal const int InvalidArgument = 3;
    internal const int PasswordRequired = 4;
    internal const int InvalidPassword = 5;
    internal const int Unsupported = 6;
    internal const int NotFound = 7;
    internal const int NotInitialized = 8;

    /// <summary>Maps an engine exception to a status code; the message travels separately as text the host may show.</summary>
    internal static int StatusOf(Exception error) => error switch {
        OperationCanceledException => Cancelled,
        PdfPasswordRequiredException => PasswordRequired,
        PdfInvalidPasswordException => InvalidPassword,
        PdfUnsupportedEncryptionException => Unsupported,
        FileNotFoundException or DirectoryNotFoundException => NotFound,
        NotSupportedException => Unsupported,
        ArgumentException => InvalidArgument,
        ObjectDisposedException => InvalidArgument,
        StudioNativeNotInitializedException => NotInitialized,
        _ => Failed
    };

    /// <summary>Copies text into memory the host releases with <c>officeimo_free</c>.</summary>
    internal static unsafe byte* ToUtf8(string value) {
        int length = Encoding.UTF8.GetByteCount(value);
        byte* buffer = (byte*)NativeMemory.Alloc((nuint)length + 1);
        fixed (char* source = value) Encoding.UTF8.GetBytes(source, value.Length, buffer, length);
        buffer[length] = 0;
        return buffer;
    }

    internal static unsafe byte* ToNative(ReadOnlySpan<byte> bytes) {
        byte* buffer = (byte*)NativeMemory.Alloc((nuint)Math.Max(1, bytes.Length));
        bytes.CopyTo(new Span<byte>(buffer, bytes.Length));
        return buffer;
    }

    internal static unsafe string? FromUtf8(byte* value) =>
        value is null ? null : Marshal.PtrToStringUTF8((nint)value);
}

/// <summary>Source-generated JSON for payloads crossing the C interface, so NativeAOT needs no reflection.</summary>
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(StudioNativeDocumentInfo))]
[JsonSerializable(typeof(StudioNativeSearchResult))]
[JsonSerializable(typeof(StudioNativeAnnotationRequest))]
[JsonSerializable(typeof(StudioNativeAnnotationList))]
[JsonSerializable(typeof(StudioNativeAnnotationUpdate))]
[JsonSerializable(typeof(StudioNativePageText))]
internal sealed partial class StudioNativeJson : JsonSerializerContext;
