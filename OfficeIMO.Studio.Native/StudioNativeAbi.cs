using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Serialization;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Conversion;
using OfficeIMO.Studio.Features.Workspace;

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
        PdfImportPasswordException { Invalid: true } => InvalidPassword,
        PdfImportPasswordException => PasswordRequired,
        PdfPasswordRequiredException => PasswordRequired,
        PdfInvalidPasswordException => InvalidPassword,
        StudioNativeCertificatePasswordException => InvalidPassword,
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
[JsonSerializable(typeof(StudioNativeSearchPage))]
[JsonSerializable(typeof(StudioNativeProtectionRequest))]
[JsonSerializable(typeof(StudioNativeSigningRequest))]
[JsonSerializable(typeof(StudioNativeAnnotationRequest))]
[JsonSerializable(typeof(StudioNativeAnnotationList))]
[JsonSerializable(typeof(StudioNativeAnnotationUpdate))]
[JsonSerializable(typeof(StudioNativePageText))]
[JsonSerializable(typeof(StudioNativeScanRequest))]
[JsonSerializable(typeof(StudioNativeScanResult))]
[JsonSerializable(typeof(StudioNativeImportRequest))]
[JsonSerializable(typeof(StudioNativeImportResult))]
[JsonSerializable(typeof(StudioNativeImportPassword))]
[JsonSerializable(typeof(IReadOnlyList<StudioNativeScanWord>))]
[JsonSerializable(typeof(StudioNativeSearchableResult))]
[JsonSerializable(typeof(StudioNativeConversionSources))]
[JsonSerializable(typeof(StudioNativeConversionResult))]
[JsonSerializable(typeof(StudioNativeFormFields))]
[JsonSerializable(typeof(StudioNativeFormFill))]
[JsonSerializable(typeof(StudioNativeStructure))]
[JsonSerializable(typeof(StudioNativeEditRequest))]
[JsonSerializable(typeof(StudioNativeWriteRequest))]
[JsonSerializable(typeof(StudioNativeWriteResult))]
internal sealed partial class StudioNativeJson : JsonSerializerContext;

internal sealed record StudioNativeScanWord(string? Text, double X, double Y, double Width, double Height, double? Confidence = null);

internal sealed record StudioNativeScanPage(string? Image, IReadOnlyList<StudioNativeScanWord>? Words = null);

internal sealed record StudioNativeScanRequest(IReadOnlyList<StudioNativeScanPage>? Pages);

internal sealed record StudioNativeScanResult(int AbiVersion, int Pages, int SearchablePages, int Words);

internal sealed record StudioNativeImportSource(string? Path, string? Password = null, IReadOnlyList<int>? Pages = null);

internal sealed record StudioNativeImportRequest(IReadOnlyList<StudioNativeImportSource>? Sources, int InsertBefore);

internal sealed record StudioNativeImportResult(int AbiVersion, int Imported, long Revision);

internal sealed record StudioNativeImportPassword(int AbiVersion, int Source, string Name);

internal sealed record StudioNativeConversionSources(int AbiVersion, IReadOnlyList<PdfConversionSource> Sources);

internal sealed record StudioNativeConversionResult(int AbiVersion, string Route, string Source, long OutputBytes, IReadOnlyList<PdfConversionNote> Notes);
