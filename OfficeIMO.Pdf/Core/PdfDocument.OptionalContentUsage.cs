using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Pdf;

internal readonly struct PdfOptionalContentUsageSummary {
    internal PdfOptionalContentUsageSummary(int pagesWithUsage, bool isComplete) {
        PagesWithUsage = pagesWithUsage;
        IsComplete = isComplete;
    }

    internal int PagesWithUsage { get; }

    internal bool IsComplete { get; }
}

public sealed partial class PdfDocument {
    internal PdfOptionalContentUsageSummary InspectPagesForOptionalContentUsage(
        IReadOnlyList<int> pageNumbers,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        PdfReadDocument document = GetReadDocument(ReadOptions, cancellationToken);
        var inspected = new HashSet<int>();
        int count = 0;
        bool isComplete = true;
        for (int index = 0; index < pageNumbers.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = pageNumbers[index];
            if (pageNumber < 1 || pageNumber > document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(pageNumbers), pageNumber, "Page number is outside the source PDF.");
            }
            if (!inspected.Add(pageNumber)) continue;
            try {
                if (document.Pages[pageNumber - 1].HasOptionalContentUsage(cancellationToken)) count++;
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception exception) when (exception is not OutOfMemoryException && exception is not StackOverflowException) {
                isComplete = false;
            }
        }
        return new PdfOptionalContentUsageSummary(count, isComplete);
    }
}
