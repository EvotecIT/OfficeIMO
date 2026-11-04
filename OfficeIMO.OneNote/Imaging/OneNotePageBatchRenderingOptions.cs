using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.OneNote;

/// <summary>Selection and rendering options for section- and notebook-level page image export.</summary>
public sealed class OneNotePageBatchRenderingOptions : OneNotePageRenderingOptions {
    /// <summary>Zero-based first page index in the flattened page sequence.</summary>
    public int PageIndex { get; set; }

    /// <summary>Maximum number of pages to export, or all remaining pages when absent.</summary>
    public int? PageCount { get; set; }

    internal OneNotePageBatchRenderingOptions CloneBatch() {
        OneNotePageBatchRenderingOptions clone = CopyTo(new OneNotePageBatchRenderingOptions());
        clone.PageIndex = PageIndex;
        clone.PageCount = PageCount;
        return clone;
    }

    internal void ValidateBatch() {
        Validate();
        if (PageIndex < 0) throw new ArgumentOutOfRangeException(nameof(PageIndex));
        if (PageCount.HasValue && PageCount.Value < 1) throw new ArgumentOutOfRangeException(nameof(PageCount));
    }
}
