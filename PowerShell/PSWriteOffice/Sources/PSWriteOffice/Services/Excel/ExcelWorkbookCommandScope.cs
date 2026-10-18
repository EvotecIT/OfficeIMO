using System;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal sealed class ExcelWorkbookCommandScope : IDisposable
{
    private bool _disposed;

    public ExcelWorkbookCommandScope(ExcelDocument document, bool ownsDocument)
    {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        OwnsDocument = ownsDocument;
    }

    public ExcelDocument Document { get; }

    public bool OwnsDocument { get; }

    public void SaveIfOwned()
    {
        if (OwnsDocument)
        {
            Document.Save();
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        if (OwnsDocument)
        {
            ExcelDocumentService.CloseDocument(Document);
        }
    }
}
