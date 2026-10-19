using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.Pdf;

internal static class PdfCommandUtilities
{
    internal static string ResolvePath(PSCmdlet cmdlet, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("Path cannot be empty.", nameof(path));
        }

        var providerPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(cmdlet.SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    internal static string ResolveExistingFilePath(PSCmdlet cmdlet, string path)
    {
        var resolvedPath = ResolvePath(cmdlet, path);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        return resolvedPath;
    }

    internal static string ResolveOutputFilePath(
        PSCmdlet cmdlet,
        string path,
        string requiredExtension,
        bool overwrite)
    {
        var resolvedPath = ResolvePath(cmdlet, path);
        if (!string.Equals(Path.GetExtension(resolvedPath), requiredExtension, StringComparison.OrdinalIgnoreCase))
        {
            throw new PSArgumentException($"OutputPath must use the {requiredExtension} extension.", nameof(path));
        }

        if (File.Exists(resolvedPath) && !overwrite)
        {
            throw new PdfOutputFileExistsException($"File '{resolvedPath}' already exists. Use -Force to overwrite it.");
        }

        return resolvedPath;
    }

    internal static void EnsureDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    internal static void EnsureOutputDirectory(string directory)
    {
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    internal static bool ShouldWrite(PSCmdlet cmdlet, string path, string action)
    {
        return cmdlet.ShouldProcess(path, action);
    }

    internal static void SetVariable(PSCmdlet cmdlet, string? name, object? value)
    {
        if (!string.IsNullOrWhiteSpace(name))
        {
            cmdlet.SessionState.PSVariable.Set(name!, value);
        }
    }

    internal static string GetSafeFileName(string fileName)
    {
        return AtomicFileWriter.GetSafeFileName(fileName, "attachment.bin");
    }

    internal static string GetUniquePath(string directory, string fileName)
    {
        var safeName = GetSafeFileName(fileName);
        var path = Path.Combine(directory, safeName);
        if (!File.Exists(path))
        {
            return path;
        }

        var extension = Path.GetExtension(safeName);
        var stem = Path.GetFileNameWithoutExtension(safeName);
        for (var index = 2; ; index++)
        {
            path = Path.Combine(directory, AtomicFileWriter.CreateCollisionCandidateName(stem, extension, index));
            if (!File.Exists(path))
            {
                return path;
            }
        }
    }

    internal static PdfDocument? ComposeContent(
        PSCmdlet cmdlet,
        PdfDocument? document,
        string parameterSetName,
        string documentParameterSet,
        Action<PdfContentBuilder> action)
    {
        if (parameterSetName == documentParameterSet)
        {
            var target = document ?? throw new PSArgumentNullException(nameof(document));
            target.Compose(compose => compose.Content(action));
            return target;
        }

        RejectDeferredPassThru(cmdlet);
        PdfDslContext.Require(cmdlet).AddContent(action);
        return null;
    }

    internal static PdfDocument? ComposePage(
        PSCmdlet cmdlet,
        PdfDocument? document,
        string parameterSetName,
        string documentParameterSet,
        Action<PdfPageBuilder> action)
    {
        if (parameterSetName == documentParameterSet)
        {
            throw new PSNotSupportedException(
                $"{cmdlet.MyInvocation.InvocationName} defines page composition and cannot update an already-built PdfDocument. Use it inside New-OfficePdf {{ ... }}.");
        }

        RejectDeferredPassThru(cmdlet);
        PdfDslContext.Require(cmdlet).ConfigurePage(action);
        return null;
    }

    internal static PdfDocument? ConfigureDocument(
        PSCmdlet cmdlet,
        PdfDocument? document,
        string parameterSetName,
        string documentParameterSet,
        Action<PdfDocument> action)
    {
        if (parameterSetName == documentParameterSet)
        {
            var target = document ?? throw new PSArgumentNullException(nameof(document));
            action(target);
            return target;
        }

        RejectDeferredPassThru(cmdlet);
        PdfDslContext.Require(cmdlet).ConfigureDocument(action);
        return null;
    }

    internal static void ConfigureOptions(PSCmdlet cmdlet, Action<PdfOptions> action)
    {
        RejectDeferredPassThru(cmdlet);
        var context = PdfDslContext.Require(cmdlet);
        action(context.Options);
    }

    private static void RejectDeferredPassThru(PSCmdlet cmdlet)
    {
        if (!cmdlet.MyInvocation.BoundParameters.TryGetValue("PassThru", out object? value) ||
            !LanguagePrimitives.IsTrue(value))
        {
            return;
        }

        throw new PSNotSupportedException(
            $"{cmdlet.MyInvocation.InvocationName} cannot emit a PdfDocument before New-OfficePdf finishes composition. Remove -PassThru inside the PDF DSL and use the document returned by New-OfficePdf.");
    }

    internal static PdfLoadOptions? CreateReadOptions(string? password, bool ignorePermissionRestrictions = false)
    {
        return CreateReadOptions(null, password, ignorePermissionRestrictions);
    }

    internal static PdfLoadOptions? CreateReadOptions(
        PdfLoadOptions? readOptions,
        string? password,
        bool ignorePermissionRestrictions = false)
    {
        if (readOptions == null && string.IsNullOrEmpty(password) && !ignorePermissionRestrictions)
        {
            return null;
        }

        var effective = readOptions ?? PdfLoadOptions.Default;
        return new PdfLoadOptions
        {
            ParsingMode = effective.ParsingMode,
            Limits = effective.Limits,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            Password = string.IsNullOrEmpty(password) ? effective.Password : password,
            PermissionPolicy = ignorePermissionRestrictions
                ? PdfPermissionPolicy.IgnoreRestrictions
                : effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    /// <summary>Loads a fluent PDF through OfficeIMO's bounded, single-snapshot path API.</summary>
    internal static PdfDocument LoadDocument(string path, PdfLoadOptions? readOptions = null)
    {
        return PdfDocument.Load(path, readOptions);
    }

    internal static ErrorRecord CreateConversionErrorRecord(
        Exception exception,
        string errorId,
        object targetObject,
        bool outputOperation)
    {
        var category = exception switch
        {
            FileNotFoundException => ErrorCategory.ObjectNotFound,
            PdfOutputFileExistsException => ErrorCategory.ResourceExists,
            PSArgumentException => ErrorCategory.InvalidArgument,
            PdfPasswordRequiredException => ErrorCategory.SecurityError,
            PdfInvalidPasswordException => ErrorCategory.SecurityError,
            PdfPermissionDeniedException => ErrorCategory.PermissionDenied,
            UnauthorizedAccessException => ErrorCategory.PermissionDenied,
            PdfReadLimitException => ErrorCategory.LimitsExceeded,
            IOException => outputOperation ? ErrorCategory.WriteError : ErrorCategory.ReadError,
            _ => ErrorCategory.InvalidOperation
        };
        return new ErrorRecord(exception, errorId, category, targetObject);
    }

    /// <summary>Opens a PDF stream only after enforcing the configured input-byte budget.</summary>
    internal static FileStream OpenBoundedReadStream(string path, PdfLoadOptions? readOptions = null)
    {
        var fullPath = Path.GetFullPath(path);
        var maxInputBytes = (readOptions?.Limits ?? new PdfReadLimits()).MaxInputBytes;
        if (maxInputBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(PdfReadLimits.MaxInputBytes), maxInputBytes, "Maximum input bytes must be positive.");
        }

        var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > maxInputBytes)
        {
            stream.Dispose();
            throw new InvalidDataException($"PDF input exceeds the configured limit of {maxInputBytes.ToString(CultureInfo.InvariantCulture)} bytes.");
        }
        if (stream.Length > int.MaxValue)
        {
            stream.Dispose();
            throw new InvalidDataException("PDF input is too large to load into a contiguous byte array.");
        }

        return stream;
    }

    internal static PdfFormFillerOptions? CreateFormFillerOptions(PSCmdlet cmdlet, string? appearanceFontPath, string? appearanceFontFamilyName, bool keepNeedAppearances)
    {
        if (string.IsNullOrWhiteSpace(appearanceFontPath) && !keepNeedAppearances)
        {
            return null;
        }

        var options = new PdfFormFillerOptions
        {
            KeepNeedAppearances = keepNeedAppearances
        };

        if (!string.IsNullOrWhiteSpace(appearanceFontPath))
        {
            var fontPath = ResolvePath(cmdlet, appearanceFontPath!);
            var familyName = string.IsNullOrWhiteSpace(appearanceFontFamilyName)
                ? Path.GetFileNameWithoutExtension(fontPath)
                : appearanceFontFamilyName!;
            options.UseAppearanceFontFile(familyName, fontPath);
        }

        return options;
    }

    internal static void ApplyEncryption(PdfOptions options, string? password, string? ownerPassword, int? permissions)
    {
        if (string.IsNullOrEmpty(password))
        {
            if (!string.IsNullOrEmpty(ownerPassword) || permissions.HasValue)
            {
                throw new PSArgumentException("-OwnerPassword and -Permission require -Password.");
            }

            return;
        }

        options.SetEncryption(password!, ownerPassword, permissions ?? PdfStandardEncryptionOptions.AllowAllPermissions);
    }

    internal static PdfDocument ApplyEncryption(PdfDocument document, string? password, string? ownerPassword, int? permissions)
    {
        if (string.IsNullOrEmpty(password))
        {
            if (!string.IsNullOrEmpty(ownerPassword) || permissions.HasValue)
            {
                throw new PSArgumentException("-OwnerPassword and -Permission require -Password.");
            }

            return document;
        }

        var options = new PdfStandardEncryptionOptions(password!)
        {
            OwnerPassword = ownerPassword,
            Permissions = permissions ?? PdfStandardEncryptionOptions.AllowAllPermissions
        };
        return document.Security.Encrypt(options).ToDocument();
    }

    internal static PdfColor? ParseColor(string? color)
        => OfficeColorUtilities.ToPdfColor(color);

    internal static PageSize ResolvePageSize(string? pageSize, double? width, double? height, bool landscape)
    {
        var pageSizeName = (pageSize ?? "Letter").Trim();
        PageSize size = pageSizeName.ToUpperInvariant() switch
        {
            "CUSTOM" => width.HasValue && height.HasValue
                ? new PageSize(width.Value, height.Value)
                : throw new PSArgumentException("Custom page size requires -Width and -Height."),
            _ => PageSizes.TryGet(pageSizeName, out var resolved)
                ? resolved
                : throw new PSArgumentException("PageSize must be a known OfficeIMO page size name or Custom.", nameof(pageSize))
        };

        return landscape ? size.Landscape() : size.Portrait();
    }

    internal static PdfPageResizeOptions? CreatePageResizeOptions(string? pageSize, double? width, double? height, bool landscape, PdfPageResizeMode mode, double? margin, bool requested)
    {
        if (!requested)
        {
            return null;
        }

        var resolvedPageSize = ResolvePageSize(pageSize ?? (width.HasValue || height.HasValue ? "Custom" : "Letter"), width, height, landscape);
        var options = new PdfPageResizeOptions(resolvedPageSize)
        {
            Mode = mode
        };

        if (margin.HasValue)
        {
            options.Margin = margin.Value;
        }

        return options;
    }

    internal static string[][] ConvertToTableRows(
        object[] inputObject,
        string[]? property,
        string[]? header,
        string collectionSeparator = ", ")
        => ConvertToTableRows(
            inputObject,
            property,
            header,
            CreateTableNormalizationOptions(collectionSeparator));

    internal static string[][] ConvertToTableRows(
        object[] inputObject,
        string[]? property,
        string[]? header,
        PowerShellObjectNormalizerOptions normalizationOptions)
    {
        if (inputObject.Length == 0)
        {
            throw new PSArgumentException("Provide at least one input object.", nameof(inputObject));
        }

        var propertyNames = property != null && property.Length > 0
            ? property
            : GetPropertyNames(inputObject[0], normalizationOptions);

        var rows = new List<string[]>();
        if (header != null && header.Length > 0)
        {
            rows.Add(header);
        }
        else
        {
            rows.Add(propertyNames);
        }

        foreach (var item in inputObject)
        {
            if (PowerShellObjectNormalizer.TryProjectItem(
                    item,
                    propertyNames,
                    out _,
                    out var values,
                    normalizationOptions))
            {
                rows.Add(values
                    .Select(value => PowerShellObjectNormalizer.NormalizeCellText(value, normalizationOptions))
                    .ToArray());
                continue;
            }

            rows.Add(new[] { PowerShellObjectNormalizer.NormalizeCellText(item, normalizationOptions) });
        }

        return rows.ToArray();
    }

    internal static string[][] ConvertDataRows(
        IEnumerable rows,
        string[]? header = null,
        string collectionSeparator = ", ")
        => ConvertDataRows(rows, header, CreateTableNormalizationOptions(collectionSeparator));

    internal static string[][] ConvertDataRows(
        IEnumerable rows,
        string[]? header,
        PowerShellObjectNormalizerOptions normalizationOptions)
    {
        var result = new List<string[]>();
        if (header != null && header.Length > 0)
        {
            result.Add(header);
        }

        foreach (var row in rows)
        {
            if (row is string[] strings)
            {
                result.Add(strings);
                continue;
            }

            if (row is IEnumerable enumerable && row is not string)
            {
                result.Add(enumerable
                    .Cast<object?>()
                    .Select(value => PowerShellObjectNormalizer.NormalizeCellText(value, normalizationOptions))
                    .ToArray());
                continue;
            }

            result.Add(new[] { PowerShellObjectNormalizer.NormalizeCellText(row, normalizationOptions) });
        }

        if (result.Count == 0)
        {
            throw new PSArgumentException("Provide at least one table row.", nameof(rows));
        }

        return result.ToArray();
    }

    internal static PowerShellObjectNormalizerOptions CreateTableNormalizationOptions(
        string collectionSeparator,
        string dictionaryEntrySeparator = "; ",
        string dictionaryKeyValueSeparator = ": ",
        int maxCollectionItems = PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems,
        int maxNestingDepth = PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth)
    {
        if (collectionSeparator == null)
        {
            throw new PSArgumentNullException(nameof(collectionSeparator));
        }

        var options = PowerShellObjectNormalizerOptions.ForTable(
            collectionSeparator,
            dictionaryEntrySeparator,
            dictionaryKeyValueSeparator,
            maxCollectionItems,
            maxNestingDepth);
        options.Culture = CultureInfo.InvariantCulture;
        options.FormatScalarValuesAsText = true;
        return options;
    }

    internal static IReadOnlyDictionary<string, string> ConvertFieldValues(IDictionary fieldValues)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (DictionaryEntry entry in fieldValues)
        {
            var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(key))
            {
                throw new PSArgumentException("Form field names cannot be empty.", nameof(fieldValues));
            }

            result[key!] = Convert.ToString(entry.Value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        return result;
    }

    internal static void ApplyPageRange(PdfTextStampOptions options, string? pageRange)
    {
        if (!string.IsNullOrWhiteSpace(pageRange))
        {
            options.UsePageRanges(PdfPageRange.ParseMany(pageRange!));
        }
    }

    internal static void ApplyPageRange(PdfImageStampOptions options, string? pageRange)
    {
        if (!string.IsNullOrWhiteSpace(pageRange))
        {
            options.UsePageRanges(PdfPageRange.ParseMany(pageRange!));
        }
    }

    private static string[] GetPropertyNames(
        object item,
        PowerShellObjectNormalizerOptions normalizationOptions)
    {
        if (PowerShellObjectNormalizer.TryProjectItem(
                item,
                columns: null,
                out var columns,
                out _,
                normalizationOptions))
        {
            return columns;
        }

        return PSObject.AsPSObject(item).Properties
            .Where(property => property.IsGettable)
            .Select(property => property.Name)
            .ToArray();
    }

    private static bool TryGetDictionaryValue(IDictionary dictionary, string key, out object? value)
    {
        if (dictionary.Contains(key))
        {
            value = dictionary[key];
            return true;
        }

        foreach (DictionaryEntry entry in dictionary)
        {
            if (string.Equals(Convert.ToString(entry.Key, CultureInfo.InvariantCulture), key, StringComparison.OrdinalIgnoreCase))
            {
                value = entry.Value;
                return true;
            }
        }

        value = null;
        return false;
    }
}

internal sealed class PdfOutputFileExistsException : IOException
{
    internal PdfOutputFileExistsException(string message) : base(message)
    {
    }
}
