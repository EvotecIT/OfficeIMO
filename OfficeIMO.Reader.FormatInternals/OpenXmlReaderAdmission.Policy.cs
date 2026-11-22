namespace OfficeIMO.Reader.FormatInternals;

// Shared with adapters that open the package without the Open XML SDK.
internal static partial class OpenXmlReaderAdmission {
    internal static OfficePackageSecurityOptions? Policy(ReaderOptions options, long nativeInputLimit, bool requirePolicy = false) {
        if (!requirePolicy && !options.MaxInputBytes.HasValue && !options.InputByteBudget.HasValue
            && !options.OpenXmlMaxPartUncompressedBytes.HasValue && !options.OpenXmlMaxTotalUncompressedBytes.HasValue) return null;
        var policy = new OfficePackageSecurityOptions { MaxPackageBytes = nativeInputLimit };
        // Reader's input guard handles zero-byte admission; the native policy requires a positive ceiling.
        if (options.MaxInputBytes.HasValue) policy.MaxPackageBytes = Math.Min(policy.MaxPackageBytes, Math.Max(1, options.MaxInputBytes.Value));
        if (options.InputByteBudget.HasValue) policy.MaxPackageBytes = Math.Min(policy.MaxPackageBytes, Math.Max(1, options.InputByteBudget.Value));
        if (options.OpenXmlMaxPartUncompressedBytes.HasValue) policy.MaxPartUncompressedBytes = Math.Min(policy.MaxPartUncompressedBytes, options.OpenXmlMaxPartUncompressedBytes.Value);
        if (options.OpenXmlMaxTotalUncompressedBytes.HasValue) policy.MaxTotalUncompressedBytes = Math.Min(policy.MaxTotalUncompressedBytes, options.OpenXmlMaxTotalUncompressedBytes.Value);
        return policy;
    }

    internal static T Load<T>(ReaderOptions options, Func<T> load) {
        try { return load(); }
        catch (OfficeInspectionResourceLimitException exception) {
            throw NativeReaderAdmission.TranslateLimit(exception);
        } catch (OfficePackageSecurityException exception) when (exception.Rule == OfficePackageSecurityRule.PartSize && options.OpenXmlMaxPartUncompressedBytes.HasValue) {
            throw new ReaderResourceLimitException(nameof(options.OpenXmlMaxPartUncompressedBytes),
                Math.Min(new OfficePackageSecurityOptions().MaxPartUncompressedBytes, options.OpenXmlMaxPartUncompressedBytes!.Value), exception);
        } catch (OfficePackageSecurityException exception) when (exception.Rule == OfficePackageSecurityRule.TotalUncompressedSize && options.OpenXmlMaxTotalUncompressedBytes.HasValue) {
            throw new ReaderResourceLimitException(nameof(options.OpenXmlMaxTotalUncompressedBytes),
                Math.Min(new OfficePackageSecurityOptions().MaxTotalUncompressedBytes, options.OpenXmlMaxTotalUncompressedBytes!.Value), exception);
        }
    }
}
