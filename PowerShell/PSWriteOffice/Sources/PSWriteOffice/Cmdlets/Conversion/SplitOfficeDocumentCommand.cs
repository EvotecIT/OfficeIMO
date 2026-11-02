using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Splits PDF files into consecutive groups, selected ranges or top-level bookmarked chapters.</summary>
/// <para>Each completed file is returned immediately. Native page extraction retains its permission,
/// signature and preservation policies. Only the selected files are written; other folder contents remain intact.</para>
/// <example>
/// <summary>Split each incoming scan into two-page records.</summary>
/// <code>Get-ChildItem ./Scans -Filter *.pdf | Split-OfficeDocument -OutputDirectory ./Records -PageCount 2 -NoClobber</code>
/// </example>
/// <example>
/// <summary>Keep nested sections inside their chapter.</summary>
/// <code>Split-OfficeDocument -LiteralPath './Manual [final].pdf' -OutputDirectory ./Chapters -ByBookmark</code>
/// </example>
[Cmdlet(VerbsCommon.Split, "OfficeDocument", DefaultParameterSetName = "PathByPageCount", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SplitOfficeDocumentCommand : AsyncPSCmdlet {
    private readonly List<string> _paths = new();
    private PdfPageSelection[]? _selections;
    private PdfLoadOptions _load = new();
    private string _destination = string.Empty;

    /// <summary>Input files in provider and pipeline order.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathByPageCount")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathByRange")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathByBookmark")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal files; also binds FileInfo.FullName and provider PSPath properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathByPageCount")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathByRange")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathByBookmark")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Destination folder. File names begin with Prefix or the source base name.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [ValidateNotNullOrEmpty]
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>Consecutive source pages per output; defaults to one.</summary>
    [Parameter(ParameterSetName = "PathByPageCount")]
    [Parameter(ParameterSetName = "LiteralPathByPageCount")]
    [ValidateRange(1, int.MaxValue)]
    public int PageCount { get; set; } = 1;

    /// <summary>One output per range string, for example '1-3','4,6'.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathByRange")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathByRange")]
    [ValidateNotNullOrEmpty]
    public string[]? Page { get; set; }

    /// <summary>Split at top-level bookmarks; nested pages remain within their chapter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathByBookmark")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathByBookmark")]
    public SwitchParameter ByBookmark { get; set; }

    /// <summary>Only these exact top-level bookmark titles; omission selects every chapter.</summary>
    [Parameter(ParameterSetName = "PathByBookmark")]
    [Parameter(ParameterSetName = "LiteralPathByBookmark")]
    [ValidateNotNullOrEmpty]
    public string[]? Bookmark { get; set; }

    /// <summary>Native PDF load settings or property overrides; includes resource and permission policies.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public object[]? Options { get; set; }

    /// <summary>Input password; accepts a credential or warned plain text.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Portable output file-name prefix; defaults to the source base name.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? Prefix { get; set; }

    /// <summary>Pad numeric part indexes to the width of the source page count.</summary>
    [Parameter]
    public SwitchParameter PadIndex { get; set; }

    /// <summary>Explicit numeric index width, up to 32 characters.</summary>
    [Parameter]
    [ValidateRange(1, 32)]
    public int IndexWidth { get; set; }

    /// <summary>Refuse to replace an existing part, including a file created by another writer.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Maximum concurrent input files.</summary>
    [Parameter]
    [ValidateRange(1, 1024)]
    public int ThrottleLimit { get; set; } = 4;

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
        _paths.AddRange(OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal));
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override Task BeginProcessingAsync() {
        bool bookmarks = ParameterSetName.EndsWith("ByBookmark", StringComparison.Ordinal);
        if (bookmarks && !ByBookmark.IsPresent) throw new PSArgumentException("ByBookmark must be enabled.");
        _selections = Page?.Select(PdfPageSelection.Parse).ToArray();
        if (Prefix != null && (Prefix != OfficePortableFileName.SanitizeBaseName(Prefix) || string.IsNullOrWhiteSpace(Prefix)))
            throw new PSArgumentException("Prefix must be a portable file-name base without directory components.");
        _destination = OfficePath.ResolveOutput(SessionState, OutputDirectory, directory: true);
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        var load = new PdfLoadOptions();
        foreach (object value in Options ?? Array.Empty<object>()) {
            object? option = OfficeArgumentValue.Unwrap(value);
            if (option is PdfLoadOptions typed) load = typed.Clone();
            else if (option is IDictionary table)
                load = ((PdfLoadOptions)new OfficeOptionsTransformationAttribute(typeof(PdfLoadOptions)).ApplyTo(engine, load, table)).Clone();
            else throw new PSArgumentException("Options requires PDF load options or a property hashtable.");
        }
        if (Password != null) load = load.Clone(null, new NetworkCredential(string.Empty, Password).Password);
        _load = load;
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync() {
        bool bookmarks = ParameterSetName.EndsWith("ByBookmark", StringComparison.Ordinal);
        PdfPageSelection[]? selections = _selections;
        PdfLoadOptions load = _load;
        string destination = _destination;
        string? prefix = Prefix;
        int pageCount = PageCount, width = IndexWidth;
        bool pad = PadIndex, noClobber = NoClobber;
        string[] titles = Bookmark?.ToArray() ?? Array.Empty<string>();
        string[] sources = _paths.ToArray();
        string[] admitted = _paths.Where(path => {
            if (OfficeConverter.FormatForExtension(System.IO.Path.GetExtension(path)) != OfficeDocumentFormat.Pdf) {
                WriteError(OfficeError.Create(new NotSupportedException("Native splitting currently supports PDF files."), "split-unsupported", path));
                return false;
            }
            return ShouldProcess(destination, "Split " + path);
        }).ToArray();
        await OfficeBatch.RunAsync(admitted, async (source, token) => {
            using PdfDocument document = await PdfDocument.LoadAsync(source, load.Clone(), token).ConfigureAwait(false);
            int count = document.Inspect(null, token).PageCount;
            PdfPageSelection[] groups;
            if (bookmarks) {
                var chapters = document.Pages.BookmarkPageRanges(1, token, titles);
                if (chapters.Count == 0) throw new InvalidOperationException("The PDF has no readable top-level chapter bookmarks.");
                groups = chapters.Select(chapter => PdfPageSelection.FromRanges(chapter.PageRange)).ToArray();
            } else if (selections != null) groups = selections;
            else {
                var ranges = new List<PdfPageSelection>();
                for (int first = 1; first <= count;) {
                    int last = (int)Math.Min(count, (long)first + pageCount - 1);
                    ranges.Add(PdfPageSelection.FromRanges(PdfPageRange.From(first, last)));
                    if (last == count) break;
                    first = last + 1;
                }
                groups = ranges.ToArray();
            }
            foreach (var group in groups) { token.ThrowIfCancellationRequested(); group.Resolve(count); }
            string stem = prefix ?? OfficePortableFileName.SanitizeBaseName(System.IO.Path.GetFileNameWithoutExtension(source));
            if (string.IsNullOrWhiteSpace(stem)) stem = "document";
            int digits = width > 0 ? width : pad ? count.ToString(CultureInfo.InvariantCulture).Length : 0;
            string[] outputs = Enumerable.Range(1, groups.Length).Select(index => System.IO.Path.Combine(destination,
                stem + "-" + index.ToString(digits > 0 ? "D" + digits.ToString(CultureInfo.InvariantCulture) : "D", CultureInfo.InvariantCulture) + ".pdf")).ToArray();
            foreach (string output in outputs) {
                if (sources.Any(input => OfficePath.AreEquivalent(input, output)))
                    throw new IOException("A split output would overwrite an input document: " + output);
                OfficePath.ValidateDestination(output, noClobber);
            }
            for (int index = 0; index < groups.Length; index++) {
                token.ThrowIfCancellationRequested();
                using PdfDocument part = document.Pages.Extract(groups[index], token);
                await part.SaveAsync(outputs[index], new PdfSaveOptions {
                    ConflictPolicy = noClobber ? OfficeConversionFileConflictPolicy.FailIfExists : OfficeConversionFileConflictPolicy.Replace
                }, token).ConfigureAwait(false);
                WriteObject(new FileInfo(outputs[index]));
            }
            return null;
        }, ThrottleLimit, CancelToken, this, "Split documents", "split-failed");
    }
}
