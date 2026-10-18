using System;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.Markdown;

namespace PSWriteOffice.Services.Markdown;

internal sealed class MarkdownDslContext : IDisposable
{
    private static readonly AsyncLocal<MarkdownDslContext?> CurrentScope = new();
    private readonly MarkdownDslContext? _previousScope;

    private MarkdownDslContext(MarkdownDoc document, MarkdownDslContext? previousScope)
    {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        _previousScope = previousScope;
    }

    public MarkdownDoc Document { get; }

    public static MarkdownDslContext Enter(MarkdownDoc document)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        var scope = new MarkdownDslContext(document, CurrentScope.Value);
        CurrentScope.Value = scope;
        return scope;
    }

    public static MarkdownDslContext Require(PSCmdlet caller)
    {
        var scope = CurrentScope.Value;
        if (scope == null)
        {
            throw new InvalidOperationException(
                $"'{caller.MyInvocation.InvocationName}' must run inside New-OfficeMarkdown script block.");
        }

        return scope;
    }

    public void Dispose()
    {
        if (CurrentScope.Value == this)
        {
            CurrentScope.Value = _previousScope;
        }
    }
}
