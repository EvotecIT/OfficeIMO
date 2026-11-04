using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

internal sealed class PdfDslContext : IDisposable
{
    private static readonly System.Threading.AsyncLocal<PdfDslContext?> Current = new();
    private readonly PdfDslContext? _previous;
    private readonly PdfDocument _document;
    private readonly List<Action<PdfContentBuilder>> _contentActions = new();
    private readonly List<Action<PdfPageBuilder>> _pageActions = new();
    private readonly List<Action<PdfDocument>> _documentActions = new();

    private PdfDslContext(PdfDocument document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _previous = Current.Value;
        Current.Value = this;
    }

    public static PdfDslContext Enter(PdfDocument document) => new(document);

    public static PdfDslContext Require(PSCmdlet cmdlet)
    {
        return Current.Value ?? throw new PSInvalidOperationException(
            $"No active PDF DSL context. Use {cmdlet.MyInvocation.InvocationName} inside New-OfficePdf {{ ... }} or pass -Document.");
    }

    public void AddContent(Action<PdfContentBuilder> action)
        => _contentActions.Add(action ?? throw new ArgumentNullException(nameof(action)));

    public void ConfigurePage(Action<PdfPageBuilder> action)
        => _pageActions.Add(action ?? throw new ArgumentNullException(nameof(action)));

    public void ConfigureOptions(Action<PdfCreateOptions> action)
        => _document.Compose(compose => compose.Settings(action));

    public void ConfigureDocument(Action<PdfDocument> action)
    {
        if (action == null) throw new ArgumentNullException(nameof(action));
        _documentActions.Add(action);
    }

    public void Apply()
    {
        _document.Compose(compose =>
        {
            if (_pageActions.Count == 0)
            {
                compose.Content(ApplyContent);
                return;
            }

            compose.Page(page =>
            {
                foreach (var action in _pageActions)
                {
                    action(page);
                }

                page.Content(content => content.Item(ApplyContent));
            });
        });

        foreach (var action in _documentActions)
        {
            action(_document);
        }
    }

    private void ApplyContent(PdfContentBuilder content)
    {
        foreach (var action in _contentActions)
        {
            action(content);
        }
    }

    public void Dispose()
    {
        Current.Value = _previous;
    }
}
