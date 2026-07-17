using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Detects a file kind from extension and bounded content evidence.
    /// </summary>
    public static ReaderDetectionResult Detect(string path, ReaderDetectionOptions? options = null) {
        ValidateFilePath(path);
        ReaderDetectionOptions effectiveOptions = NormalizeDetectionOptions(options);
        ReaderDetectionResult extensionResult = BuildExtensionDetection(path);
        if (!ShouldInspectContent(extensionResult.ExtensionKind, effectiveOptions.Mode)) {
            return extensionResult;
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return DetectCore(stream, path, effectiveOptions, extensionResult);
    }

    /// <summary>
    /// Detects a stream kind from source-name and bounded content evidence. Seekable streams are restored
    /// to their original position. Non-seekable streams are advanced by at most <see cref="ReaderDetectionOptions.MaxProbeBytes"/>.
    /// </summary>
    public static ReaderDetectionResult Detect(
        Stream stream,
        string? sourceName = null,
        ReaderDetectionOptions? options = null) {
        ValidateReadableStream(stream);
        ReaderDetectionOptions effectiveOptions = NormalizeDetectionOptions(options);
        string logicalSourceName = NormalizeLogicalSourceName(sourceName, "memory");
        ReaderDetectionResult extensionResult = BuildExtensionDetection(logicalSourceName);
        if (!ShouldInspectContent(extensionResult.ExtensionKind, effectiveOptions.Mode)) {
            return extensionResult;
        }

        return DetectCore(stream, logicalSourceName, effectiveOptions, extensionResult);
    }

    /// <summary>
    /// Detects a byte payload kind from source-name and bounded content evidence.
    /// </summary>
    public static ReaderDetectionResult Detect(
        byte[] bytes,
        string? sourceName = null,
        ReaderDetectionOptions? options = null) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return Detect(stream, sourceName, options);
    }

    /// <summary>
    /// Asynchronously detects a file kind from extension and bounded content evidence.
    /// </summary>
    public static async Task<ReaderDetectionResult> DetectAsync(
        string path,
        ReaderDetectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateFilePath(path);
        ReaderDetectionOptions effectiveOptions = NormalizeDetectionOptions(options);
        ReaderDetectionResult extensionResult = BuildExtensionDetection(path);
        if (!ShouldInspectContent(extensionResult.ExtensionKind, effectiveOptions.Mode)) {
            return extensionResult;
        }

        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            bufferSize: 4096,
            useAsync: true);
        return await DetectCoreAsync(stream, path, effectiveOptions, extensionResult, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asynchronously detects a stream kind from source-name and bounded content evidence.
    /// Seekable streams are restored to their original position.
    /// </summary>
    public static async Task<ReaderDetectionResult> DetectAsync(
        Stream stream,
        string? sourceName = null,
        ReaderDetectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateReadableStream(stream);
        ReaderDetectionOptions effectiveOptions = NormalizeDetectionOptions(options);
        string logicalSourceName = NormalizeLogicalSourceName(sourceName, "memory");
        ReaderDetectionResult extensionResult = BuildExtensionDetection(logicalSourceName);
        if (!ShouldInspectContent(extensionResult.ExtensionKind, effectiveOptions.Mode)) {
            return extensionResult;
        }

        return await DetectCoreAsync(
            stream,
            logicalSourceName,
            effectiveOptions,
            extensionResult,
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asynchronously detects a byte payload kind from source-name and bounded content evidence.
    /// </summary>
    public static async Task<ReaderDetectionResult> DetectAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderDetectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return await DetectAsync(stream, sourceName, options, cancellationToken).ConfigureAwait(false);
    }

    private static ReaderDetectionResult DetectForRead(string path, ReaderOptions options) {
        return Detect(path, CreateDetectionOptions(options));
    }

    private static ReaderDetectionResult DetectForRead(Stream stream, string? sourceName, ReaderOptions options) {
        return Detect(stream, sourceName, CreateDetectionOptions(options));
    }

    private static bool TryResolvePathHandler(
        string path,
        ReaderOptions options,
        out ReaderHandlerDescriptor handler,
        out ReaderDetectionResult detection) {
        detection = DetectForRead(path, options);
        return TrySelectPathHandler(path, options, detection, out handler);
    }

    private static bool TrySelectPathHandler(
        string path,
        ReaderOptions options,
        ReaderDetectionResult detection,
        out ReaderHandlerDescriptor handler) {
        bool hasExtensionHandler = TryResolveCustomHandlerByPath(path, out ReaderHandlerDescriptor extensionHandler) &&
                                   extensionHandler.SupportsPathInput;
        bool contentOverridesExtension = options.DetectionMode == ReaderDetectionMode.PreferContent &&
                                         detection.IsMismatch;
        if (contentOverridesExtension) {
            return TryResolveDetectedHandler(detection, pathInput: true, out handler);
        }
        if (hasExtensionHandler) {
            handler = extensionHandler;
            return true;
        }
        return TryResolveDetectedHandler(detection, pathInput: true, out handler);
    }

    private static bool TryResolveStreamHandler(
        Stream stream,
        string? sourceName,
        ReaderOptions options,
        out ReaderHandlerDescriptor handler,
        out ReaderDetectionResult detection) {
        detection = DetectForRead(stream, sourceName, options);
        return TrySelectStreamHandler(sourceName, options, detection, out handler);
    }

    private static bool TrySelectStreamHandler(
        string? sourceName,
        ReaderOptions options,
        ReaderDetectionResult detection,
        out ReaderHandlerDescriptor handler) {
        bool hasExtensionHandler = TryResolveCustomHandlerBySourceName(sourceName, out ReaderHandlerDescriptor extensionHandler) &&
                                   extensionHandler.SupportsStreamInput;
        bool contentOverridesExtension = options.DetectionMode == ReaderDetectionMode.PreferContent &&
                                         detection.IsMismatch;
        if (contentOverridesExtension) {
            return TryResolveDetectedHandler(detection, pathInput: false, out handler);
        }
        if (hasExtensionHandler) {
            handler = extensionHandler;
            return true;
        }
        return TryResolveDetectedHandler(detection, pathInput: false, out handler);
    }

    private static bool TryResolveDetectedHandler(
        ReaderDetectionResult detection,
        bool pathInput,
        out ReaderHandlerDescriptor handler) {
        if (TryResolveCustomHandlerByKind(detection.Kind, pathInput, out handler)) {
            return true;
        }

        if (!CanUseZipContainerFallback(detection)) {
            return false;
        }

        return TryResolveCustomHandlerByKind(ReaderInputKind.Zip, pathInput, out handler);
    }

    private static bool CanUseZipContainerFallback(ReaderDetectionResult detection) {
        return detection.Kind == ReaderInputKind.OpenDocument;
    }

    private static async Task<HandlerDetectionResolution> ResolvePathHandlerAsync(
        string path,
        ReaderOptions options,
        CancellationToken cancellationToken) {
        ReaderDetectionResult detection = await DetectAsync(
            path,
            CreateDetectionOptions(options),
            cancellationToken).ConfigureAwait(false);
        bool hasHandler = TrySelectPathHandler(path, options, detection, out ReaderHandlerDescriptor handler);
        return new HandlerDetectionResolution(hasHandler ? handler : null, detection);
    }

    private static async Task<HandlerDetectionResolution> ResolveStreamHandlerAsync(
        Stream stream,
        string? sourceName,
        ReaderOptions options,
        CancellationToken cancellationToken) {
        ReaderDetectionResult detection = await DetectAsync(
            stream,
            sourceName,
            CreateDetectionOptions(options),
            cancellationToken).ConfigureAwait(false);
        bool hasHandler = TrySelectStreamHandler(sourceName, options, detection, out ReaderHandlerDescriptor handler);
        return new HandlerDetectionResolution(hasHandler ? handler : null, detection);
    }

    private static ReaderDetectionOptions CreateDetectionOptions(ReaderOptions options) {
        return new ReaderDetectionOptions {
            Mode = options.DetectionMode,
            MaxProbeBytes = options.DetectionMaxProbeBytes,
            MaxContainerEntries = options.DetectionMaxContainerEntries,
            InspectContainers = true
        };
    }

    private static ReaderDetectionResult DetectCore(
        Stream stream,
        string sourceName,
        ReaderDetectionOptions options,
        ReaderDetectionResult extensionResult) {
        long originalPosition = 0;
        bool restorePosition = stream.CanSeek;
        if (restorePosition) {
            originalPosition = stream.Position;
        }

        try {
            byte[] prefix = ReadDetectionPrefix(stream, options.MaxProbeBytes);
            DetectionCandidate candidate = DetectContent(prefix);
            bool containerInspected = false;

            if (candidate.Kind == ReaderInputKind.Zip && options.InspectContainers && restorePosition) {
                DetectionCandidate containerCandidate = InspectZipContainer(
                    stream,
                    originalPosition,
                    options.MaxContainerEntries);
                containerInspected = true;
                if (containerCandidate.Kind != ReaderInputKind.Unknown) {
                    candidate = containerCandidate;
                }
            } else if (IsOleCompound(prefix) && options.InspectContainers) {
                candidate = restorePosition
                    ? InspectEmailCompound(stream, originalPosition, options.MaxContainerEntries)
                    : InspectEmailCompound(prefix);
                containerInspected = true;
            } else if (IsCabinet(prefix) && options.InspectContainers) {
                candidate = restorePosition
                    ? InspectCabinetContainer(stream, originalPosition, options.MaxContainerEntries)
                    : InspectCabinetContainer(prefix, options.MaxContainerEntries);
                containerInspected = true;
            }

            return BuildCombinedDetection(
                extensionResult,
                candidate,
                prefix.Length,
                containerInspected,
                options.Mode);
        } finally {
            if (restorePosition) {
                stream.Position = originalPosition;
            }
        }
    }

    private static async Task<ReaderDetectionResult> DetectCoreAsync(
        Stream stream,
        string sourceName,
        ReaderDetectionOptions options,
        ReaderDetectionResult extensionResult,
        CancellationToken cancellationToken) {
        long originalPosition = 0;
        bool restorePosition = stream.CanSeek;
        if (restorePosition) originalPosition = stream.Position;

        try {
            byte[] prefix = await ReadDetectionPrefixAsync(stream, options.MaxProbeBytes, cancellationToken).ConfigureAwait(false);
            DetectionCandidate candidate = DetectContent(prefix);
            bool containerInspected = false;
            if (candidate.Kind == ReaderInputKind.Zip && options.InspectContainers && restorePosition) {
                DetectionCandidate containerCandidate = await InspectZipContainerAsync(
                    stream,
                    originalPosition,
                    options.MaxContainerEntries,
                    cancellationToken).ConfigureAwait(false);
                containerInspected = true;
                if (containerCandidate.Kind != ReaderInputKind.Unknown) candidate = containerCandidate;
            } else if (IsOleCompound(prefix) && options.InspectContainers) {
                candidate = restorePosition
                    ? InspectEmailCompound(stream, originalPosition, options.MaxContainerEntries)
                    : InspectEmailCompound(prefix);
                containerInspected = true;
            } else if (IsCabinet(prefix) && options.InspectContainers) {
                candidate = restorePosition
                    ? await InspectCabinetContainerAsync(
                        stream,
                        originalPosition,
                        options.MaxContainerEntries,
                        cancellationToken).ConfigureAwait(false)
                    : InspectCabinetContainer(prefix, options.MaxContainerEntries);
                containerInspected = true;
            }

            return BuildCombinedDetection(extensionResult, candidate, prefix.Length, containerInspected, options.Mode);
        } finally {
            if (restorePosition) stream.Position = originalPosition;
        }
    }

    private static ReaderDetectionResult BuildExtensionDetection(string sourceName) {
        string extension = NormalizeExtension(TryGetExtension(sourceName));
        ReaderInputKind extensionKind = extension.Length == 0 ? ReaderInputKind.Unknown : DetectKind(sourceName);
        var evidence = extension.Length == 0
            ? Array.Empty<string>()
            : new[] { "extension:" + extension };

        return new ReaderDetectionResult {
            SourceName = sourceName,
            Extension = extension,
            ExtensionKind = extensionKind,
            ExtensionConfidence = extensionKind == ReaderInputKind.Unknown
                ? ReaderDetectionConfidence.None
                : ReaderDetectionConfidence.Medium,
            Kind = extensionKind,
            Confidence = extensionKind == ReaderInputKind.Unknown
                ? ReaderDetectionConfidence.None
                : ReaderDetectionConfidence.Medium,
            MediaType = GetMediaType(extension, extensionKind),
            Evidence = evidence
        };
    }

    private static ReaderDetectionResult BuildCombinedDetection(
        ReaderDetectionResult extensionResult,
        DetectionCandidate content,
        int inspectedBytes,
        bool containerInspected,
        ReaderDetectionMode mode) {
        var evidence = new List<string>(extensionResult.Evidence);
        evidence.AddRange(content.Evidence);

        ReaderInputKind effectiveKind = extensionResult.ExtensionKind;
        ReaderDetectionConfidence effectiveConfidence = extensionResult.ExtensionConfidence;
        if (extensionResult.ExtensionKind == content.Kind &&
            content.Kind != ReaderInputKind.Unknown &&
            content.Confidence >= ReaderDetectionConfidence.Medium) {
            effectiveConfidence = ReaderDetectionConfidence.High;
        } else if (mode == ReaderDetectionMode.PreferContent &&
            content.Kind != ReaderInputKind.Unknown &&
            content.Confidence >= ReaderDetectionConfidence.Medium) {
            effectiveKind = content.Kind;
            effectiveConfidence = content.Confidence;
        } else if (extensionResult.ExtensionKind == ReaderInputKind.Unknown && content.Kind != ReaderInputKind.Unknown) {
            effectiveKind = content.Kind;
            effectiveConfidence = content.Confidence;
        }

        bool useContentMediaType = content.Kind == effectiveKind &&
            content.MediaType != null &&
            (effectiveKind != extensionResult.ExtensionKind ||
             (mode == ReaderDetectionMode.PreferContent && content.MediaTypeIsDeclared));
        string? effectiveMediaType = useContentMediaType
            ? content.MediaType
            : effectiveKind == extensionResult.ExtensionKind
                ? extensionResult.MediaType ?? GetMediaType(effectiveKind)
                : GetMediaType(effectiveKind);

        return new ReaderDetectionResult {
            SourceName = extensionResult.SourceName,
            Extension = extensionResult.Extension,
            ExtensionKind = extensionResult.ExtensionKind,
            ContentKind = content.Kind,
            ExtensionConfidence = extensionResult.ExtensionConfidence,
            ContentConfidence = content.Confidence,
            Kind = effectiveKind,
            Confidence = effectiveConfidence,
            MediaType = effectiveMediaType,
            ContentInspected = true,
            ContainerInspected = containerInspected,
            InspectedBytes = inspectedBytes,
            Evidence = evidence.ToArray()
        };
    }

    private static bool ShouldInspectContent(ReaderInputKind extensionKind, ReaderDetectionMode mode) {
        return mode == ReaderDetectionMode.PreferContent ||
               (mode == ReaderDetectionMode.ContentWhenUnknown && extensionKind == ReaderInputKind.Unknown);
    }

    private static ReaderDetectionOptions NormalizeDetectionOptions(ReaderDetectionOptions? options) {
        ReaderDetectionOptions source = options ?? new ReaderDetectionOptions();
        return new ReaderDetectionOptions {
            Mode = Enum.IsDefined(typeof(ReaderDetectionMode), source.Mode)
                ? source.Mode
                : ReaderDetectionMode.PreferContent,
            MaxProbeBytes = Math.Min(ReaderOptions.MaximumDetectionProbeBytes, Math.Max(256, source.MaxProbeBytes)),
            MaxContainerEntries = Math.Min(ReaderOptions.MaximumDetectionContainerEntries, Math.Max(1, source.MaxContainerEntries)),
            InspectContainers = source.InspectContainers
        };
    }

    private static byte[] ReadDetectionPrefix(Stream stream, int maxBytes) {
        var buffer = new byte[maxBytes];
        int total = 0;
        while (total < buffer.Length) {
            int read = stream.Read(buffer, total, buffer.Length - total);
            if (read <= 0) break;
            total += read;
        }

        if (total == buffer.Length) return buffer;
        var exact = new byte[total];
        Array.Copy(buffer, exact, total);
        return exact;
    }

    private static async Task<byte[]> ReadDetectionPrefixAsync(
        Stream stream,
        int maxBytes,
        CancellationToken cancellationToken) {
        var buffer = new byte[maxBytes];
        int total = 0;
        while (total < buffer.Length) {
            int read = await stream.ReadAsync(buffer, total, buffer.Length - total, cancellationToken).ConfigureAwait(false);
            if (read <= 0) break;
            total += read;
        }

        if (total == buffer.Length) return buffer;
        var exact = new byte[total];
        Array.Copy(buffer, exact, total);
        return exact;
    }

    private static DetectionCandidate DetectContent(byte[] prefix) {
        if (prefix.Length == 0) {
            return DetectionCandidate.Unknown("content:empty");
        }
        if (StartsWith(prefix, new byte[] { 0xE4, 0x52, 0x5C, 0x7B, 0x8C, 0xD8, 0xA7, 0x4D, 0xAE, 0xB1, 0x53, 0x78, 0xD0, 0x29, 0x96, 0xD3 })) {
            return DetectionCandidate.High(ReaderInputKind.OneNote, "application/onenote", "signature:onenote-section");
        }
        if (StartsWith(prefix, new byte[] { 0xA1, 0x2F, 0xFF, 0x43, 0xD9, 0xEF, 0x76, 0x4C, 0x9E, 0xE2, 0x10, 0xEA, 0x57, 0x22, 0x76, 0x5F })) {
            return DetectionCandidate.High(ReaderInputKind.OneNote, "application/onenote", "signature:onenote-toc");
        }
        if (IndexOfAscii(prefix, "%PDF-", Math.Min(prefix.Length, 1024)) >= 0) {
            return DetectionCandidate.High(ReaderInputKind.Pdf, "application/pdf", "signature:pdf");
        }
        if (StartsWith(prefix, new byte[] { 0x50, 0x4B, 0x03, 0x04 }) ||
            StartsWith(prefix, new byte[] { 0x50, 0x4B, 0x05, 0x06 }) ||
            StartsWith(prefix, new byte[] { 0x50, 0x4B, 0x07, 0x08 })) {
            return DetectionCandidate.High(ReaderInputKind.Zip, "application/zip", "signature:zip");
        }
        if (StartsWith(prefix, new byte[] { 0x78, 0x9F, 0x3E, 0x22 })) {
            return DetectionCandidate.High(ReaderInputKind.Email, "application/ms-tnef", "signature:tnef");
        }
        if (StartsWith(prefix, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 })) {
            return DetectionCandidate.Unknown("signature:ole-compound");
        }
        if (IsCabinet(prefix)) {
            return DetectionCandidate.Unknown("signature:cabinet");
        }

        string? text = DecodeTextPrefix(prefix);
        if (text == null) {
            return DetectionCandidate.Unknown("content:binary");
        }

        string trimmed = text.TrimStart('\uFEFF', ' ', '\t', '\r', '\n');
        string lower = trimmed.Length > 4096 ? trimmed.Substring(0, 4096).ToLowerInvariant() : trimmed.ToLowerInvariant();
        if (StartsWithContentLineRoot(lower, "begin:vcalendar")) {
            return DetectionCandidate.High(ReaderInputKind.Calendar, "text/calendar", "text:icalendar-root");
        }
        if (StartsWithContentLineRoot(lower, "begin:vcard")) {
            return DetectionCandidate.High(ReaderInputKind.VCard, "text/vcard", "text:vcard-root");
        }
        if (LooksLikeEmailMessage(trimmed)) {
            return DetectionCandidate.High(ReaderInputKind.Email, "message/rfc822", "text:rfc-message-headers");
        }
        if (lower.StartsWith("{\\rtf", StringComparison.Ordinal)) {
            return DetectionCandidate.High(ReaderInputKind.Rtf, "application/rtf", "signature:rtf");
        }
        if (lower.StartsWith("<!doctype html", StringComparison.Ordinal) ||
            lower.StartsWith("<html", StringComparison.Ordinal)) {
            return DetectionCandidate.High(ReaderInputKind.Html, "text/html", "text:html-root");
        }
        if (lower.StartsWith("<?xml", StringComparison.Ordinal)) {
            return DetectionCandidate.Medium(ReaderInputKind.Xml, "application/xml", "text:xml-declaration");
        }
        if (trimmed.StartsWith("{", StringComparison.Ordinal) || trimmed.StartsWith("[", StringComparison.Ordinal)) {
            return DetectionCandidate.Medium(ReaderInputKind.Json, "application/json", "text:json-leading-token");
        }
        if (LooksLikeMarkdown(trimmed)) {
            return DetectionCandidate.Medium(ReaderInputKind.Markdown, "text/markdown", "text:markdown-marker");
        }
        if (trimmed.StartsWith("<", StringComparison.Ordinal)) {
            return DetectionCandidate.Low(ReaderInputKind.Xml, "application/xml", "text:xml-root");
        }

        return DetectionCandidate.Low(ReaderInputKind.Text, "text/plain", "content:mostly-text");
    }

    private static bool StartsWithContentLineRoot(string value, string root) {
        if (!value.StartsWith(root, StringComparison.Ordinal)) return false;
        return value.Length == root.Length || value[root.Length] == '\r' || value[root.Length] == '\n';
    }

    private static bool LooksLikeEmailMessage(string text) {
        string normalized = text.Replace("\r\n", "\n");
        string[] lines = normalized.Split('\n');
        int recognizedHeaders = 0;
        bool hasAddressHeader = false;
        bool hasMessageHeader = false;

        for (int index = 0; index < lines.Length && index < 256; index++) {
            string line = lines[index];
            if (index == 0 && line.StartsWith("From ", StringComparison.Ordinal)) {
                hasMessageHeader = true;
                continue;
            }
            if (line.Length == 0) break;
            if ((line[0] == ' ' || line[0] == '\t') && recognizedHeaders > 0) continue;

            int colon = line.IndexOf(':');
            if (colon <= 0 || colon > 64) return false;
            string name = line.Substring(0, colon).Trim();
            if (string.Equals(name, "From", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Sender", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "To", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Cc", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Bcc", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Reply-To", StringComparison.OrdinalIgnoreCase)) {
                recognizedHeaders++;
                hasAddressHeader = true;
            } else if (string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Subject", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Message-ID", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "MIME-Version", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Content-Type", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Content-Transfer-Encoding", StringComparison.OrdinalIgnoreCase)) {
                recognizedHeaders++;
                hasMessageHeader = true;
            }
        }

        return recognizedHeaders >= 2 && hasAddressHeader && hasMessageHeader;
    }

    private static DetectionCandidate? MatchContainerEntry(string name) {
        return name switch {
            "word/document.xml" => DetectionCandidate.High(ReaderInputKind.Word, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "container:word/document.xml"),
            "xl/workbook.xml" => DetectionCandidate.High(ReaderInputKind.Excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "container:xl/workbook.xml"),
            "ppt/presentation.xml" => DetectionCandidate.High(ReaderInputKind.PowerPoint, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "container:ppt/presentation.xml"),
            "visio/document.xml" => DetectionCandidate.High(ReaderInputKind.Visio, "application/vnd.ms-visio.drawing.main+xml", "container:visio/document.xml"),
            _ => null
        };
    }

    private static bool LooksLikeMarkdown(string text) {
        if (text.StartsWith("# ", StringComparison.Ordinal) ||
            text.StartsWith("## ", StringComparison.Ordinal) ||
            text.StartsWith("```", StringComparison.Ordinal)) {
            return true;
        }

        return text.IndexOf("\n# ", StringComparison.Ordinal) >= 0 ||
               text.IndexOf("\n## ", StringComparison.Ordinal) >= 0 ||
               text.IndexOf("\n```", StringComparison.Ordinal) >= 0;
    }

    private static string? DecodeTextPrefix(byte[] bytes) {
        if (StartsWith(bytes, new byte[] { 0xFF, 0xFE })) return Encoding.Unicode.GetString(bytes);
        if (StartsWith(bytes, new byte[] { 0xFE, 0xFF })) return Encoding.BigEndianUnicode.GetString(bytes);

        int zeroCount = 0;
        int controlCount = 0;
        for (int index = 0; index < bytes.Length; index++) {
            byte value = bytes[index];
            if (value == 0) zeroCount++;
            if (value < 0x09 || (value > 0x0D && value < 0x20)) controlCount++;
        }
        if (zeroCount > bytes.Length / 8 || controlCount > bytes.Length / 8) {
            return null;
        }

        return Encoding.UTF8.GetString(bytes);
    }

    private static bool StartsWith(byte[] source, byte[] signature) {
        if (source.Length < signature.Length) return false;
        for (int index = 0; index < signature.Length; index++) {
            if (source[index] != signature[index]) return false;
        }
        return true;
    }

    private static int IndexOfAscii(byte[] source, string value, int limit) {
        byte[] pattern = Encoding.ASCII.GetBytes(value);
        int last = Math.Min(source.Length, limit) - pattern.Length;
        for (int index = 0; index <= last; index++) {
            int patternIndex = 0;
            while (patternIndex < pattern.Length && source[index + patternIndex] == pattern[patternIndex]) {
                patternIndex++;
            }
            if (patternIndex == pattern.Length) return index;
        }
        return -1;
    }

    private static bool ReadExact(Stream stream, byte[] buffer, int offset, int count) {
        int total = 0;
        while (total < count) {
            int read = stream.Read(buffer, offset + total, count - total);
            if (read <= 0) return false;
            total += read;
        }
        return true;
    }

    private static async Task<bool> ReadExactAsync(
        Stream stream,
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        int total = 0;
        while (total < count) {
            int read = await stream.ReadAsync(buffer, offset + total, count - total, cancellationToken).ConfigureAwait(false);
            if (read <= 0) return false;
            total += read;
        }
        return true;
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) {
        return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
    }

    private static uint ReadUInt32(byte[] bytes, int offset) {
        return (uint)(bytes[offset] |
                      (bytes[offset + 1] << 8) |
                      (bytes[offset + 2] << 16) |
                      (bytes[offset + 3] << 24));
    }

    private static string? GetMediaType(ReaderInputKind kind) {
        return kind switch {
            ReaderInputKind.Word => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ReaderInputKind.Excel => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ReaderInputKind.PowerPoint => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ReaderInputKind.Markdown => "text/markdown",
            ReaderInputKind.Pdf => "application/pdf",
            ReaderInputKind.Email => "message/rfc822",
            ReaderInputKind.Calendar => "text/calendar",
            ReaderInputKind.VCard => "text/vcard",
            ReaderInputKind.OneNote => "application/onenote",
            ReaderInputKind.Text => "text/plain",
            ReaderInputKind.Csv => "text/csv",
            ReaderInputKind.Json => "application/json",
            ReaderInputKind.Xml => "application/xml",
            ReaderInputKind.Html => "text/html",
            ReaderInputKind.Zip => "application/zip",
            ReaderInputKind.Epub => "application/epub+zip",
            ReaderInputKind.Visio => "application/vnd.ms-visio.drawing.main+xml",
            ReaderInputKind.Yaml => "application/yaml",
            ReaderInputKind.Rtf => "application/rtf",
            _ => null
        };
    }

    private static string? GetMediaType(string extension, ReaderInputKind kind) {
        return extension switch {
            ".docm" => "application/vnd.ms-word.document.macroEnabled.12",
            ".doc" => "application/msword",
            ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
            ".xls" => "application/vnd.ms-excel",
            ".pptm" => "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
            ".eml" => "message/rfc822",
            ".msg" => "application/vnd.ms-outlook",
            ".oft" => "application/vnd.ms-outlook",
            ".mbox" or ".mbx" => "application/mbox",
            ".tnef" => "application/ms-tnef",
            ".ics" => "text/calendar",
            ".vcs" => "text/x-vcalendar",
            ".vcf" or ".vcard" => "text/vcard",
            ".csv" => "text/csv",
            ".tsv" => "text/tab-separated-values",
            ".json" => "application/json",
            ".xml" => "application/xml",
            ".yml" or ".yaml" => "application/yaml",
            ".odt" => "application/vnd.oasis.opendocument.text",
            ".ods" => "application/vnd.oasis.opendocument.spreadsheet",
            ".odp" => "application/vnd.oasis.opendocument.presentation",
            _ => GetMediaType(kind)
        };
    }

    private static ReaderInputKind NormalizeBuiltInDispatchKind(ReaderInputKind kind) {
        return kind switch {
            ReaderInputKind.Csv or
            ReaderInputKind.Json or
            ReaderInputKind.Xml or
            ReaderInputKind.Html or
            ReaderInputKind.Yaml or
            ReaderInputKind.Rtf => ReaderInputKind.Text,
            ReaderInputKind.Word or
            ReaderInputKind.Excel or
            ReaderInputKind.PowerPoint or
            ReaderInputKind.Markdown or
            ReaderInputKind.Pdf or
            ReaderInputKind.Email or
            ReaderInputKind.Calendar or
            ReaderInputKind.VCard or
            ReaderInputKind.Text => kind,
            _ => ReaderInputKind.Unknown
        };
    }

    private sealed class DetectionCandidate {
        private DetectionCandidate(
            ReaderInputKind kind,
            ReaderDetectionConfidence confidence,
            string? mediaType,
            IReadOnlyList<string> evidence,
            bool mediaTypeIsDeclared = false) {
            Kind = kind;
            Confidence = confidence;
            MediaType = mediaType;
            Evidence = evidence;
            MediaTypeIsDeclared = mediaTypeIsDeclared;
        }

        public ReaderInputKind Kind { get; }
        public ReaderDetectionConfidence Confidence { get; }
        public string? MediaType { get; }
        public IReadOnlyList<string> Evidence { get; }
        public bool MediaTypeIsDeclared { get; }

        public static DetectionCandidate Unknown(string evidence) =>
            new DetectionCandidate(ReaderInputKind.Unknown, ReaderDetectionConfidence.None, null, new[] { evidence });

        public static DetectionCandidate Low(ReaderInputKind kind, string mediaType, string evidence) =>
            new DetectionCandidate(kind, ReaderDetectionConfidence.Low, mediaType, new[] { evidence });

        public static DetectionCandidate Medium(ReaderInputKind kind, string mediaType, string evidence) =>
            new DetectionCandidate(kind, ReaderDetectionConfidence.Medium, mediaType, new[] { evidence });

        public static DetectionCandidate High(
            ReaderInputKind kind,
            string mediaType,
            string evidence,
            bool mediaTypeIsDeclared = false) =>
            new DetectionCandidate(
                kind,
                ReaderDetectionConfidence.High,
                mediaType,
                new[] { evidence },
                mediaTypeIsDeclared);
    }

    private sealed class HandlerDetectionResolution {
        public HandlerDetectionResolution(ReaderHandlerDescriptor? handler, ReaderDetectionResult detection) {
            Handler = handler;
            Detection = detection ?? throw new ArgumentNullException(nameof(detection));
        }

        public ReaderHandlerDescriptor? Handler { get; }
        public ReaderDetectionResult Detection { get; }
    }
}
