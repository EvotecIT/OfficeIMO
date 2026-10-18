using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Splits a PDF into page, range, count, or bookmark files.</summary>
/// <example>
///   <summary>Split a PDF into page files.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pages = Split-OfficePdf -Path .\Examples\Documents\Combined.pdf -OutputDirectory .\Examples\Documents\Pages -Prefix 'combined-page'
/// $pages | Select-Object Name, Length</code>
///   <para>Creates one output PDF for each page and returns the written files.</para>
/// </example>
[Cmdlet(VerbsCommon.Split, "OfficePdf", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SplitOfficePdfCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output directory.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>Output file prefix.</summary>
    [Parameter]
    public string Prefix { get; set; } = "page";

    /// <summary>Create one PDF for each consecutive group with this many pages.</summary>
    [Parameter]
    public int PagesPerDocument { get; set; }

    /// <summary>Create one PDF for each supplied page range or selection, such as 1-3 or 1,3.</summary>
    [Parameter]
    [Alias("Range", "PageRanges")]
    public string[]? PageRange { get; set; }

    /// <summary>Create one PDF for each supplied bookmark title.</summary>
    [Parameter]
    [Alias("Bookmark", "BookmarkTitle")]
    public string[]? BookmarkName { get; set; }

    /// <summary>Create one PDF for every readable bookmark when -BookmarkName is not supplied.</summary>
    [Parameter]
    public SwitchParameter ByBookmark { get; set; }

    /// <summary>Password used to open an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed assembly restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Pad numeric split names to the width required by the source page count.</summary>
    [Parameter]
    public SwitchParameter PadIndex { get; set; }

    /// <summary>Pad numeric split names to this explicit width.</summary>
    [Parameter]
    public int? IndexWidth { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PdfPageSelection[]? pageSelections = ValidateArguments();
        var outputDirectory = PdfCommandUtilities.ResolvePath(this, OutputDirectory);
        var document = PdfDocument.Load(
            PdfCommandUtilities.ResolvePath(this, Path),
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
        var outputs = CreateOutputs(document, pageSelections);
        foreach (var output in outputs)
        {
            var outputPath = PdfCommandUtilities.GetUniquePath(outputDirectory, Prefix + "-" + output.Name + ".pdf");
            if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write split PDF"))
            {
                continue;
            }

            PdfCommandUtilities.EnsureOutputDirectory(outputDirectory);
            output.Document
                .SaveAsync(outputPath, OfficeConversionFileConflictPolicy.FailIfExists, System.Threading.CancellationToken.None)
                .GetAwaiter()
                .GetResult()
                .RequireSuccess();
            WriteObject(new FileInfo(outputPath));
        }
    }

    private IReadOnlyList<PdfSplitOutput> CreateOutputs(
        PdfDocument document,
        PdfPageSelection[]? pageSelections)
    {
        var pageCount = document.Inspect().PageCount;
        var indexWidth = GetIndexWidth(pageCount);

        if (PagesPerDocument > 0)
        {
            var ranges = GetPageGroups(pageCount, PagesPerDocument);
            var documents = document.Pages.Split(ranges);
            return CreateRangeOutputs(documents, ranges, indexWidth);
        }

        if (pageSelections is { Length: > 0 })
        {
            var documents = document.Pages.Split(pageSelections);
            return documents
                .Select((pdf, index) => new PdfSplitOutput(pdf, PdfCommandUtilities.GetSafeFileName(PageRange![index])))
                .ToArray();
        }

        if (ByBookmark.IsPresent || BookmarkName is { Length: > 0 })
        {
            var names = BookmarkName ?? Array.Empty<string>();
            var ranges = document.Pages.BookmarkPageRanges(names);
            var documents = document.Pages.Split(ranges.Select(range => range.PageRange));
            return documents
                .Select((pdf, index) => new PdfSplitOutput(pdf, PdfCommandUtilities.GetSafeFileName(ranges[index].Title)))
                .ToArray();
        }

        var split = document.Pages.Split();
        return split
            .Select((pdf, index) => new PdfSplitOutput(pdf, FormatIndex(index + 1, indexWidth)))
            .ToArray();
    }

    private PdfPageSelection[]? ValidateArguments()
    {
        var modes = 0;
        if (PagesPerDocument > 0)
        {
            modes++;
        }

        if (PageRange != null && PageRange.Length > 0)
        {
            modes++;
        }

        if (ByBookmark.IsPresent || BookmarkName is { Length: > 0 })
        {
            modes++;
        }

        if (modes > 1)
        {
            throw new PSArgumentException("Use only one split mode: -PagesPerDocument, -PageRange, or -ByBookmark/-BookmarkName.");
        }

        if (PagesPerDocument < 0)
        {
            throw new PSArgumentException("-PagesPerDocument must be greater than zero.", nameof(PagesPerDocument));
        }

        if (IndexWidth.HasValue && IndexWidth.Value < 1)
        {
            throw new PSArgumentException("-IndexWidth must be greater than zero.", nameof(IndexWidth));
        }

        return PageRange is { Length: > 0 }
            ? PageRange.Select(PdfPageSelection.Parse).ToArray()
            : null;
    }

    private IReadOnlyList<PdfSplitOutput> CreateRangeOutputs(IReadOnlyList<PdfDocument> documents, IReadOnlyList<PdfPageRange> ranges, int indexWidth)
    {
        return documents
            .Select((pdf, index) => new PdfSplitOutput(pdf, FormatRange(ranges[index], indexWidth)))
            .ToArray();
    }

    private int GetIndexWidth(int pageCount)
    {
        if (IndexWidth.HasValue)
        {
            return IndexWidth.Value;
        }

        return PadIndex.IsPresent
            ? Math.Max(1, pageCount.ToString(CultureInfo.InvariantCulture).Length)
            : 0;
    }

    private string FormatRange(PdfPageRange range, int indexWidth) =>
        range.FirstPage == range.LastPage
            ? FormatIndex(range.FirstPage, indexWidth)
            : FormatIndex(range.FirstPage, indexWidth) + "-" + FormatIndex(range.LastPage, indexWidth);

    private static string FormatIndex(int index, int indexWidth) =>
        indexWidth > 0
            ? index.ToString("D" + indexWidth.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
            : index.ToString(CultureInfo.InvariantCulture);

    private static IReadOnlyList<PdfPageRange> GetPageGroups(int pageCount, int pagesPerDocument)
    {
        if (pageCount <= 0)
        {
            throw new PSInvalidOperationException("PDF does not contain any readable pages.");
        }

        var ranges = new List<PdfPageRange>();
        for (var firstPage = 1; firstPage <= pageCount; firstPage += pagesPerDocument)
        {
            var lastPage = Math.Min(firstPage + pagesPerDocument - 1, pageCount);
            ranges.Add(PdfPageRange.From(firstPage, lastPage));
        }

        return ranges;
    }

    private sealed class PdfSplitOutput
    {
        internal PdfSplitOutput(PdfDocument document, string name)
        {
            Document = document;
            Name = name;
        }

        internal PdfDocument Document { get; }

        internal string Name { get; }
    }
}
