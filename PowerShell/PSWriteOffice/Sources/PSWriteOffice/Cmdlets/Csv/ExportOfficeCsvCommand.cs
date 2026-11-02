using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Exports objects or a CSV document to a CSV file.</summary>
/// <para>Use <c>ConvertTo-OfficeCsv</c> when the destination should be CSV text in the pipeline.</para>
/// <example>
///   <summary>Export objects to a CSV file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data | Export-OfficeCsv -Path .\export.csv</code>
///   <para>Streams PowerShell objects into a CSV file.</para>
/// </example>
/// <example>
///   <summary>Export with culture list separator.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data | Export-OfficeCsv -Path .\export.csv -UseCulture -Culture pl-PL</code>
///   <para>Uses the selected culture list separator as the delimiter.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeCsv", DefaultParameterSetName = ParameterSetInputObjectPathDelimiter, SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed partial class ExportOfficeCsvCommand : PSCmdlet
{
    private const int StreamWriterBufferSize = 64 * 1024;
    private const int CompressionProbeBufferSize = 256;
    private const string ParameterSetInputObjectPathDelimiter = "InputObjectPathDelimiter";
    private const string ParameterSetInputObjectPathCulture = "InputObjectPathCulture";
    private const string ParameterSetInputObjectLiteralPathDelimiter = "InputObjectLiteralPathDelimiter";
    private const string ParameterSetInputObjectLiteralPathCulture = "InputObjectLiteralPathCulture";
    private const string ParameterSetDocumentPathDelimiter = "DocumentPathDelimiter";
    private const string ParameterSetDocumentPathCulture = "DocumentPathCulture";
    private const string ParameterSetDocumentLiteralPathDelimiter = "DocumentLiteralPathDelimiter";
    private const string ParameterSetDocumentLiteralPathCulture = "DocumentLiteralPathCulture";
    private CsvRowWriter? _streamingWriter;
    private readonly CsvPowerShellObjectProjector _objectProjector = new();
    private string? _resolvedPath;
    private string[]? _appendHeader;
    private Encoding? _appendEncoding;
    private bool _appendToExistingFile;
    private bool _skipOutput;
    private bool _wroteOutput;

    /// <summary>Objects to export into CSV rows.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectLiteralPathCulture)]
    public object? InputObject { get; set; }

    /// <summary>CSV document to export.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPathCulture)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentLiteralPathCulture)]
    public CsvDocument Document { get; set; } = null!;

    /// <summary>Destination CSV path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentPathCulture)]
    [Alias("FilePath", "OutputPath", "OutPath")]
    public string? Path { get; set; }

    /// <summary>Literal destination CSV path. Wildcards are not expanded.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectLiteralPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentLiteralPathCulture)]
    [Alias("PSPath", "LP")]
    public string? LiteralPath { get; set; }

    /// <summary>Field delimiter character.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentLiteralPathDelimiter)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Field delimiter text for multi-character delimiters such as || or ::.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentLiteralPathDelimiter)]
    public string? DelimiterText { get; set; }

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectLiteralPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentLiteralPathCulture)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Omit the header row from the output.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Override the newline sequence.</summary>
    [Parameter]
    public string? NewLine { get; set; }

    /// <summary>Culture used for value formatting.</summary>
    [Parameter]
    public CultureInfo? Culture { get; set; }

    /// <summary>Encoding used when writing files.</summary>
    [Parameter]
    public Encoding? Encoding { get; set; }

    /// <summary>Compression used when writing the CSV file.</summary>
    [Parameter]
    public CsvCompressionType CompressionType { get; set; } = CsvCompressionType.Auto;

    /// <summary>Compression level used when <see cref="CompressionType"/> is enabled.</summary>
    [Parameter]
    public CompressionLevel CompressionLevel { get; set; } = CompressionLevel.Optimal;

    /// <summary>Controls how formula-like values are written.</summary>
    [Parameter]
    public CsvFormulaInjectionPolicy FormulaInjectionPolicy { get; set; } = CsvFormulaInjectionPolicy.Preserve;

    /// <summary>Controls when CSV fields are quoted. Defaults to quoting only fields that need it.</summary>
    [Parameter]
    public CsvQuoteMode UseQuotes { get; set; } = CsvQuoteMode.AsNeeded;

    /// <summary>Field names that should always be quoted when <see cref="UseQuotes"/> is AsNeeded.</summary>
    [Parameter]
    public string[]? QuoteFields { get; set; }

    /// <summary>Token written for null values.</summary>
    [Parameter]
    public string? NullValue { get; set; }

    /// <summary>Date/time format used for DateTime and DateTimeOffset values.</summary>
    [Parameter]
    public string? DateTimeFormat { get; set; }

    /// <summary>Convert date/time values to UTC before formatting.</summary>
    [Parameter]
    public SwitchParameter UseUtc { get; set; }

    /// <summary>Maximum number of items allowed in one nested collection or dictionary field. Defaults to 1,048,575; increase explicitly for trusted larger values.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectLiteralPathCulture)]
    [ValidateRange(1, int.MaxValue)]
    public int MaxCollectionItems { get; set; } = PSWriteOffice.Services.PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems;

    /// <summary>Maximum nesting depth allowed while normalizing one field value. Defaults to 64; increase explicitly for trusted deeper values.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectLiteralPathCulture)]
    [ValidateRange(1, int.MaxValue)]
    public int MaxNestingDepth { get; set; } = PSWriteOffice.Services.PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth;

    /// <summary>Emit a <see cref="FileInfo"/> for the exported file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Append rows to an existing CSV file. Existing headers are reused when present.</summary>
    [Parameter]
    public SwitchParameter Append { get; set; }

    /// <summary>Do not overwrite an existing CSV file.</summary>
    [Parameter]
    [Alias("NoOverwrite")]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Allow overwriting read-only files and appending rows with missing existing columns.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        if (UseCulture.IsPresent)
        {
            var separator = (Culture ?? CultureInfo.CurrentCulture).TextInfo.ListSeparator;
            if (string.IsNullOrEmpty(separator) || separator.Length != 1)
            {
                throw new PSArgumentException("The selected culture must provide a single-character list separator.");
            }

            Delimiter = separator[0];
        }
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (IsDocumentParameterSet())
        {
            ExportDocument(Document);
            return;
        }

        if (TryGetCsvDocument(InputObject, out var csvDocument))
        {
            ExportDocument(csvDocument);
            return;
        }

        if (TryGetDataTable(InputObject, out var dataTable))
        {
            ExportDataTable(dataTable);
            return;
        }

        if (TryGetDataView(InputObject, out var dataView))
        {
            ExportDataTable(dataView.ToTable());
            return;
        }

        if (TryGetDataReader(InputObject, out var dataReader))
        {
            ExportDataReader(dataReader);
            return;
        }

        WriteStreamingInputObject(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (_streamingWriter != null)
        {
            DisposeStreamingWriter();
            WritePassThru();
        }
    }

    /// <inheritdoc />
    protected override void StopProcessing()
    {
        DisposeStreamingWriter();
    }

    private void ExportDocument(CsvDocument document)
    {
        if (document == null || !TryPrepareOutput("Write CSV", allowAdditionalAppend: Append.IsPresent))
        {
            return;
        }

        if (Append.IsPresent)
        {
            AppendDocument(document);
        }
        else
        {
            document.Save(_resolvedPath!, CreateSaveOptions());
        }

        _wroteOutput = true;
        WritePassThru();
    }

    private static bool TryGetCsvDocument(object? value, out CsvDocument document)
    {
        if (value is CsvDocument csvDocument)
        {
            document = csvDocument;
            return true;
        }

        if (value is PSObject { BaseObject: CsvDocument psObjectDocument })
        {
            document = psObjectDocument;
            return true;
        }

        document = null!;
        return false;
    }

    private void AppendDocument(CsvDocument document)
    {
        var options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
        var appendHeader = GetEffectiveAppendHeader(document);

        if (appendHeader is { Length: > 0 })
        {
            ValidateDocumentAppendHeader(document, appendHeader);
        }

        using var csvWriter = CreateRowWriter(append: true, options);

        if (appendHeader is { Length: > 0 })
        {
            WriteDocumentRows(document, csvWriter, appendHeader, projectByName: true);
            return;
        }

        WriteDocumentRows(document, csvWriter, document.Header, projectByName: false);
    }

    private void WriteStreamingInputObject(object? value)
    {
        var writer = EnsureStreamingWriter(value);
        if (writer == null)
        {
            return;
        }

        try
        {
            _objectProjector.WriteObject(value, writer);
        }
        catch
        {
            DisposeStreamingWriter();
            throw;
        }
    }

    private CsvRowWriter? EnsureStreamingWriter(object? firstValue)
    {
        if (_streamingWriter != null)
        {
            return _streamingWriter;
        }

        if (!TryPrepareOutput("Write CSV"))
        {
            return null;
        }

        var options = CreateSaveOptions();
        _objectProjector.UseCsvOptions(options, MaxCollectionItems, MaxNestingDepth);
        if (Append.IsPresent)
        {
            options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
            _objectProjector.UseCsvOptions(options, MaxCollectionItems, MaxNestingDepth);
            var appendHeader = GetEffectiveAppendHeader(firstValue);
            if (appendHeader is { Length: > 0 })
            {
                if (!Force.IsPresent)
                {
                    _objectProjector.ValidateObjectColumns(firstValue, appendHeader);
                }

                _objectProjector.UseColumns(appendHeader, validateColumns: !Force.IsPresent);
            }
        }

        _streamingWriter = CreateRowWriter(Append.IsPresent, options);
        _wroteOutput = true;
        return _streamingWriter;
    }

    private CsvRowWriter? EnsureStreamingWriterForColumns(
        IReadOnlyList<string> sourceColumns,
        out IReadOnlyList<string> effectiveColumns,
        Action<IReadOnlyList<string>>? validateBeforeOpen = null,
        bool allowAdditionalAppend = false,
        IReadOnlyList<string>? activeAppendColumns = null)
    {
        if (sourceColumns == null)
        {
            throw new ArgumentNullException(nameof(sourceColumns));
        }

        if (_streamingWriter != null)
        {
            effectiveColumns = _objectProjector.CurrentColumns ?? sourceColumns;
            return _streamingWriter;
        }

        if (!TryPrepareOutput("Write CSV", allowAdditionalAppend))
        {
            effectiveColumns = sourceColumns;
            return null;
        }

        var options = CreateSaveOptions();
        _objectProjector.UseCsvOptions(options, MaxCollectionItems, MaxNestingDepth);
        effectiveColumns = sourceColumns;
        var validateFollowingObjects = false;
        if (Append.IsPresent)
        {
            options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
            _objectProjector.UseCsvOptions(options, MaxCollectionItems, MaxNestingDepth);
            var appendHeader = activeAppendColumns is { Count: > 0 }
                ? activeAppendColumns
                : GetEffectiveAppendHeader(sourceColumns);
            if (appendHeader is { Count: > 0 })
            {
                effectiveColumns = appendHeader;
                validateFollowingObjects = !Force.IsPresent;
            }
        }

        validateBeforeOpen?.Invoke(effectiveColumns);
        _objectProjector.UseColumns(effectiveColumns, validateColumns: validateFollowingObjects);
        _streamingWriter = CreateRowWriter(Append.IsPresent, options);
        _wroteOutput = true;
        return _streamingWriter;
    }

    private bool TryPrepareOutput(string action, bool allowAdditionalAppend = false)
    {
        if (_skipOutput)
        {
            return false;
        }

        if (_wroteOutput && !(allowAdditionalAppend && Append.IsPresent))
        {
            WriteError(new ErrorRecord(
                new InvalidOperationException("Path can only be written once per invocation."),
                "CsvOutputAlreadyWritten",
                ErrorCategory.InvalidOperation,
                GetTargetPathForErrors()));
            _skipOutput = true;
            return false;
        }

        _resolvedPath = ResolveOutputPath();
        if (!ShouldProcess(_resolvedPath, action))
        {
            _skipOutput = true;
            return false;
        }

        var needsFileState = Append.IsPresent || NoClobber.IsPresent || Force.IsPresent;
        var fileExists = needsFileState && File.Exists(_resolvedPath);
        var appendTargetHasBytes = Append.IsPresent && fileExists && new FileInfo(_resolvedPath).Length > 0;
        if (appendTargetHasBytes && IsCompressedAppendTarget(CompressionType, _resolvedPath))
        {
            WriteError(new ErrorRecord(
                new NotSupportedException("Appending to compressed CSV files is not supported."),
                "CsvCompressedAppendNotSupported",
                ErrorCategory.NotImplemented,
                _resolvedPath));
            _skipOutput = true;
            return false;
        }

        if (fileExists && NoClobber.IsPresent && !Append.IsPresent)
        {
            WriteError(new ErrorRecord(
                new IOException($"File '{_resolvedPath}' already exists."),
                "CsvFileExistsNoClobber",
                ErrorCategory.ResourceExists,
                _resolvedPath));
            _skipOutput = true;
            return false;
        }

        var directory = System.IO.Path.GetDirectoryName(_resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (fileExists && Force.IsPresent)
        {
            var fileInfo = new FileInfo(_resolvedPath);
            if (fileInfo.IsReadOnly)
            {
                fileInfo.IsReadOnly = false;
            }
        }

        _appendEncoding = appendTargetHasBytes && Encoding == null
            ? TryDetectEncodingFromBom(_resolvedPath)
            : null;
        _appendToExistingFile = appendTargetHasBytes && ContainsCsvContent(_resolvedPath, Encoding ?? _appendEncoding);
        _appendHeader = _appendToExistingFile
            ? ReadAppendHeader(_resolvedPath)
            : null;

        return true;
    }

    private void DisposeStreamingWriter()
    {
        _streamingWriter?.Dispose();
        _streamingWriter = null;
        _appendHeader = null;
        _appendEncoding = null;
        _appendToExistingFile = false;
        _objectProjector.Reset();
    }

    private void WritePassThru()
    {
        if (PassThru.IsPresent && !string.IsNullOrWhiteSpace(_resolvedPath))
        {
            WriteObject(new FileInfo(_resolvedPath!));
        }
    }

    private CsvSaveOptions CreateSaveOptions(bool? includeHeader = null)
    {
        var options = new CsvSaveOptions
        {
            Delimiter = Delimiter,
            DelimiterText = DelimiterText,
            IncludeHeader = includeHeader ?? !NoHeader.IsPresent,
            Culture = Culture ?? CultureInfo.InvariantCulture,
            Encoding = Encoding,
            CompressionType = CompressionType,
            CompressionLevel = CompressionLevel,
            FormulaInjectionPolicy = FormulaInjectionPolicy,
            QuoteMode = UseQuotes,
            QuoteFields = QuoteFields
        };

        CsvPowerShellOptionBuilder.ApplySaveOptions(
            options,
            NullValue,
            DateTimeFormat,
            UseUtc.IsPresent,
            CompressionType,
            CompressionLevel);

        if (!string.IsNullOrEmpty(NewLine))
        {
            options.NewLine = NewLine!;
        }

        return options;
    }

    private bool IsDocumentParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetDocumentPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentPathCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentLiteralPathCulture, StringComparison.OrdinalIgnoreCase);

    private bool IsLiteralPathParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetInputObjectLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetInputObjectLiteralPathCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentLiteralPathCulture, StringComparison.OrdinalIgnoreCase);

    private string ResolveOutputPath()
    {
        var path = IsLiteralPathParameterSet()
            ? LiteralPath
            : Path;

        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("A destination path is required.");
        }

        return SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
    }

    private string? GetTargetPathForErrors() => IsLiteralPathParameterSet() ? LiteralPath : Path;

    private CsvRowWriter CreateRowWriter(bool append, CsvSaveOptions options)
    {
        var appendToContent = append && _appendToExistingFile;
        var compressionType = CsvRowWriter.ResolveCompression(options.CompressionType, _resolvedPath!);
        if (appendToContent && compressionType != CsvCompressionType.None)
        {
            throw new NotSupportedException("Appending to compressed CSV files is not supported.");
        }

        var encoding = ResolveOutputEncoding(append, options);
        options.Encoding = encoding;
        if (appendToContent)
        {
            EnsureAppendStartsOnNewRecord(_resolvedPath!, options);
        }

        options.Encoding = encoding;
        return CsvRowWriter.CreateFile(_resolvedPath!, options, append: appendToContent, bufferSize: StreamWriterBufferSize);
    }

    private static bool IsCompressedAppendTarget(CsvCompressionType requestedCompressionType, string path)
    {
        if (requestedCompressionType == CsvCompressionType.None)
        {
            return HasGZipHeader(path) ||
                HasDeflatePayload(path) ||
                HasBrotliPayload(path) ||
                HasZLibPayload(path);
        }

        return CsvRowWriter.ResolveCompression(requestedCompressionType, path) != CsvCompressionType.None ||
            HasCompressedFileExtension(path) ||
            HasGZipHeader(path) ||
            HasDeflatePayload(path) ||
            HasBrotliPayload(path) ||
            HasZLibPayload(path);
    }

    private static bool HasCompressedFileExtension(string path)
    {
        var extension = System.IO.Path.GetExtension(path);
        return string.Equals(extension, ".gz", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".gzip", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".deflate", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".br", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".brotli", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".zlib", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasGZipHeader(string path)
    {
        var header = new byte[2];
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, StreamWriterBufferSize, FileOptions.SequentialScan);
        return stream.Read(header, 0, header.Length) == header.Length && header[0] == 0x1F && header[1] == 0x8B;
    }

    private static bool HasDeflatePayload(string path)
    {
        try
        {
            var buffer = new byte[CompressionProbeBufferSize];
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, StreamWriterBufferSize, FileOptions.SequentialScan);
            if (stream.Length == 0)
            {
                return false;
            }

            using var deflateStream = new DeflateStream(stream, CompressionMode.Decompress, leaveOpen: false);
            var bytesRead = deflateStream.Read(buffer, 0, buffer.Length);
            return bytesRead > 0;
        }
        catch (InvalidDataException)
        {
            return false;
        }
        catch (IOException)
        {
            return false;
        }
    }

    private static bool HasBrotliPayload(string path)
    {
#if NET8_0_OR_GREATER
        try
        {
            return CanReadCompressedPayload(path, static stream => new BrotliStream(stream, CompressionMode.Decompress, leaveOpen: false));
        }
        catch (InvalidDataException)
        {
            return false;
        }
        catch (IOException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
#else
        return false;
#endif
    }

    private static bool HasZLibPayload(string path)
    {
#if NET8_0_OR_GREATER
        try
        {
            return CanReadCompressedPayload(path, static stream => new ZLibStream(stream, CompressionMode.Decompress, leaveOpen: false));
        }
        catch (InvalidDataException)
        {
            return false;
        }
        catch (IOException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
#else
        return false;
#endif
    }

#if NET8_0_OR_GREATER
    private static bool CanReadCompressedPayload(string path, Func<Stream, Stream> streamFactory)
    {
        var buffer = new byte[CompressionProbeBufferSize];
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, StreamWriterBufferSize, FileOptions.SequentialScan);
        if (stream.Length == 0)
        {
            return false;
        }

        using var compressedStream = streamFactory(stream);
        var bytesRead = compressedStream.Read(buffer, 0, buffer.Length);
        return bytesRead > 0;
    }
#endif

    private Encoding ResolveOutputEncoding(bool append, CsvSaveOptions options) =>
        options.Encoding ?? (append ? _appendEncoding : null) ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    private static bool ContainsCsvContent(string path, Encoding? encoding)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, StreamWriterBufferSize, FileOptions.SequentialScan);
        using var reader = new StreamReader(stream, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: StreamWriterBufferSize, leaveOpen: false);
        while (reader.Read() is var value && value != -1)
        {
            if (!char.IsWhiteSpace((char)value))
            {
                return true;
            }
        }

        return false;
    }

    private void EnsureAppendStartsOnNewRecord(string path, CsvSaveOptions options)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, StreamWriterBufferSize, FileOptions.SequentialScan);
        if (stream.Length == 0)
        {
            return;
        }

        var encoding = ResolveOutputEncoding(append: true, options);
        if (StreamEndsWithNewLine(stream, encoding))
        {
            return;
        }

        stream.Position = stream.Length;
        var newLineBytes = encoding.GetBytes(options.NewLine);
        stream.Write(newLineBytes, 0, newLineBytes.Length);
    }

    private string[] ReadAppendHeader(string path)
    {
        var options = new CsvLoadOptions
        {
            Delimiter = Delimiter,
            DelimiterText = DelimiterText,
            Encoding = Encoding ?? _appendEncoding,
            Culture = Culture ?? CultureInfo.InvariantCulture,
            CompressionType = CompressionType
        };

        return CsvDocument.Load(path, options).Header.ToArray();
    }

    private static Encoding? TryDetectEncodingFromBom(string path)
    {
        var bom = new byte[4];
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, bufferSize: 4, FileOptions.SequentialScan);
        var read = stream.Read(bom, 0, bom.Length);
        if (read >= 4 && bom[0] == 0xFF && bom[1] == 0xFE && bom[2] == 0x00 && bom[3] == 0x00)
        {
            return new UTF32Encoding(bigEndian: false, byteOrderMark: true);
        }

        if (read >= 4 && bom[0] == 0x00 && bom[1] == 0x00 && bom[2] == 0xFE && bom[3] == 0xFF)
        {
            return new UTF32Encoding(bigEndian: true, byteOrderMark: true);
        }

        if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
        {
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        }

        if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
        {
            return System.Text.Encoding.Unicode;
        }

        if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
        {
            return System.Text.Encoding.BigEndianUnicode;
        }

        return null;
    }

    private static bool StreamEndsWithNewLine(FileStream stream, Encoding encoding)
    {
        if (stream.Length == 0)
        {
            return false;
        }

        if (encoding.CodePage == System.Text.Encoding.Unicode.CodePage ||
            encoding.CodePage == System.Text.Encoding.BigEndianUnicode.CodePage)
        {
            return StreamEndsWithEncodedNewLine(stream, encoding, byteCount: 2);
        }

        if (encoding.CodePage == System.Text.Encoding.UTF32.CodePage ||
            encoding.CodePage == new UTF32Encoding(bigEndian: true, byteOrderMark: true).CodePage)
        {
            return StreamEndsWithEncodedNewLine(stream, encoding, byteCount: 4);
        }

        stream.Position = stream.Length - 1;
        var lastByte = stream.ReadByte();
        return lastByte == '\n' || lastByte == '\r';
    }

    private static bool StreamEndsWithEncodedNewLine(FileStream stream, Encoding encoding, int byteCount)
    {
        if (stream.Length < byteCount)
        {
            return false;
        }

        var buffer = new byte[byteCount];
        stream.Position = stream.Length - byteCount;
        var read = stream.Read(buffer, 0, buffer.Length);
        if (read != buffer.Length)
        {
            return false;
        }

        var value = encoding.GetString(buffer);
        return value.Length > 0 && (value[value.Length - 1] == '\n' || value[value.Length - 1] == '\r');
    }

    private void ValidateDocumentAppendHeader(CsvDocument document, IReadOnlyList<string> appendHeader)
    {
        if (Force.IsPresent)
        {
            return;
        }

        var documentHeader = new HashSet<string>(document.Header, StringComparer.OrdinalIgnoreCase);
        foreach (var column in appendHeader)
        {
            if (!documentHeader.Contains(column))
            {
                throw new CsvException($"Cannot append CSV because the document is missing the existing column '{column}'. Use -Force to append with blank values for missing columns.");
            }
        }
    }

    private string[]? GetEffectiveAppendHeader(object? firstValue)
    {
        if (_appendHeader is not { Length: > 0 })
        {
            return null;
        }

        if (!NoHeader.IsPresent || Force.IsPresent || _objectProjector.CanProjectColumns(firstValue, _appendHeader))
        {
            return _appendHeader;
        }

        return null;
    }

    private string[]? GetEffectiveAppendHeader(IReadOnlyList<string> sourceColumns)
    {
        if (_appendHeader is not { Length: > 0 })
        {
            return null;
        }

        if (!NoHeader.IsPresent ||
            Force.IsPresent ||
            ContainsColumns(sourceColumns, _appendHeader))
        {
            return _appendHeader;
        }

        return null;
    }

    private string[]? GetEffectiveAppendHeader(CsvDocument document)
    {
        if (_appendHeader is not { Length: > 0 })
        {
            return null;
        }

        if (!NoHeader.IsPresent)
        {
            return _appendHeader;
        }

        var documentHeader = new HashSet<string>(document.Header, StringComparer.OrdinalIgnoreCase);
        return _appendHeader.All(documentHeader.Contains) ? _appendHeader : null;
    }

    private static bool ContainsColumns(IReadOnlyList<string> sourceColumns, IReadOnlyList<string> requiredColumns)
    {
        var columns = new HashSet<string>(sourceColumns, StringComparer.OrdinalIgnoreCase);
        return requiredColumns.All(columns.Contains);
    }

    private static void WriteDocumentRows(CsvDocument document, CsvRowWriter writer, IReadOnlyList<string> columns, bool projectByName)
    {
        foreach (var row in document.AsEnumerable())
        {
            writer.WriteRow(
                columns,
                columns.Count,
                (Row: row, Columns: columns, ProjectByName: projectByName),
                static (state, index) => state.ProjectByName
                    ? TryGetRowValue(state.Row, state.Columns[index])
                    : index < state.Row.FieldCount ? state.Row[index] : null);
        }
    }

    private static object? TryGetRowValue(CsvRow row, string column)
    {
        try
        {
            return row[column];
        }
        catch
        {
            return null;
        }
    }
}
