using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

internal sealed class PdfDslContext : IDisposable
{
    private static readonly System.Threading.AsyncLocal<PdfDslContext?> Current = new();
    private readonly PdfDslContext? _previous;
    private readonly List<Action<PdfContentBuilder>> _contentActions = new();
    private readonly List<Action<PdfPageBuilder>> _pageActions = new();
    private readonly List<Func<PdfDocument, PdfDocument>> _documentActions = new();

    private PdfDslContext(PdfOptions options)
    {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _previous = Current.Value;
        Current.Value = this;
    }

    public PdfOptions Options { get; }

    public static PdfDslContext Enter(PdfOptions options) => new(options);

    public static PdfDslContext Require(PSCmdlet cmdlet)
    {
        return Current.Value ?? throw new PSInvalidOperationException(
            $"No active PDF DSL context. Use {cmdlet.MyInvocation.InvocationName} inside New-OfficePdf {{ ... }} or pass -Document.");
    }

    public void AddContent(Action<PdfContentBuilder> action)
        => _contentActions.Add(action ?? throw new ArgumentNullException(nameof(action)));

    public void ConfigurePage(Action<PdfPageBuilder> action)
        => _pageActions.Add(action ?? throw new ArgumentNullException(nameof(action)));

    public void ConfigureDocument(Action<PdfDocument> action)
    {
        if (action == null) throw new ArgumentNullException(nameof(action));
        _documentActions.Add(document =>
        {
            action(document);
            return document;
        });
    }

    public PdfDocument Build()
    {
        var document = PdfDocument.Create(compose =>
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
        }, Options);

        foreach (var action in _documentActions)
        {
            document = action(document);
        }

        return document;
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
