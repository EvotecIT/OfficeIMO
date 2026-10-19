using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;

namespace PSWriteOffice.Services.Visio;

internal sealed class VisioDslContext : IDisposable
{
    private static readonly AsyncLocal<VisioDslContext?> CurrentScope = new();
    private readonly Stack<VisioPage> _pages = new();
    private readonly Dictionary<VisioPage, Dictionary<string, VisioShape>> _shapesByPage = new();
    private readonly Dictionary<string, VisioStencilCatalog> _stencilCatalogsByKey = new(StringComparer.OrdinalIgnoreCase);

    private VisioDslContext(VisioDocument document)
    {
        Document = document ?? throw new ArgumentNullException(nameof(document));
    }

    internal VisioDocument Document { get; }

    internal static VisioDslContext? Current => CurrentScope.Value;

    internal static VisioDslContext Enter(VisioDocument document)
    {
        if (CurrentScope.Value != null)
        {
            throw new InvalidOperationException("A Visio DSL scope is already active on this runspace.");
        }

        var scope = new VisioDslContext(document);
        CurrentScope.Value = scope;
        return scope;
    }

    internal static VisioDslContext Require(PSCmdlet caller)
    {
        return CurrentScope.Value ?? throw new PSInvalidOperationException(
            $"'{caller.MyInvocation.InvocationName}' must run inside New-OfficeVisio.");
    }

    internal VisioPage? CurrentPage => _pages.Count == 0 ? null : _pages.Peek();

    internal VisioStencilCatalog? DefaultStencilCatalog { get; private set; }

    internal VisioPage RequirePage()
    {
        return CurrentPage ?? throw new PSInvalidOperationException("No Visio page context available. Use VisioPage first.");
    }

    internal IDisposable Push(VisioPage page)
    {
        if (page == null)
        {
            throw new ArgumentNullException(nameof(page));
        }

        _pages.Push(page);
        return new PopToken(this, page);
    }

    internal void RegisterShape(VisioPage page, string? key, VisioShape shape)
    {
        if (page == null)
        {
            throw new ArgumentNullException(nameof(page));
        }

        if (shape == null)
        {
            throw new ArgumentNullException(nameof(shape));
        }

        if (!_shapesByPage.TryGetValue(page, out var shapesByKey))
        {
            shapesByKey = new Dictionary<string, VisioShape>(StringComparer.OrdinalIgnoreCase);
            _shapesByPage[page] = shapesByKey;
        }

        if (!string.IsNullOrWhiteSpace(key))
        {
            shapesByKey[key!] = shape;
        }

        if (!string.IsNullOrWhiteSpace(shape.Name))
        {
            shapesByKey[shape.Name!] = shape;
        }

        if (!string.IsNullOrWhiteSpace(shape.Id))
        {
            shapesByKey[shape.Id] = shape;
        }
    }

    internal VisioShape ResolveShape(VisioPage page, string value)
    {
        if (page == null)
        {
            throw new ArgumentNullException(nameof(page));
        }

        if (string.IsNullOrWhiteSpace(value))
        {
            throw new PSArgumentException("Shape reference cannot be empty.", nameof(value));
        }

        if (_shapesByPage.TryGetValue(page, out var shapesByKey) &&
            shapesByKey.TryGetValue(value, out var shape))
        {
            return shape;
        }

        shape = page.AllShapes().FirstOrDefault(candidate =>
            string.Equals(candidate.Id, value, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.Name, value, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.NameU, value, StringComparison.OrdinalIgnoreCase));

        if (shape != null)
        {
            RegisterShape(page, value, shape);
            return shape;
        }

        throw new PSInvalidOperationException($"Visio shape '{value}' was not found on the active Visio page.");
    }

    internal void RegisterStencilCatalog(string? key, VisioStencilCatalog catalog, bool makeDefault)
    {
        if (catalog == null)
        {
            throw new ArgumentNullException(nameof(catalog));
        }

        if (!string.IsNullOrWhiteSpace(key))
        {
            _stencilCatalogsByKey[key!] = catalog;
        }

        if (!string.IsNullOrWhiteSpace(catalog.Name))
        {
            _stencilCatalogsByKey[catalog.Name] = catalog;
        }

        if (DefaultStencilCatalog == null || makeDefault)
        {
            DefaultStencilCatalog = catalog;
        }
    }

    internal VisioStencilCatalog ResolveStencilCatalog(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return DefaultStencilCatalog ?? OfficeIMO.Visio.Stencils.VisioStencils.All;
        }

        if (_stencilCatalogsByKey.TryGetValue(value!, out var catalog))
        {
            return catalog;
        }

        throw new PSInvalidOperationException($"Visio stencil catalog '{value}' was not found in the current DSL scope.");
    }

    private void Pop(VisioPage page)
    {
        if (_pages.Count > 0 && ReferenceEquals(_pages.Peek(), page))
        {
            _pages.Pop();
        }
    }

    public void Dispose()
    {
        if (CurrentScope.Value == this)
        {
            CurrentScope.Value = null;
        }

        _pages.Clear();
        _shapesByPage.Clear();
        _stencilCatalogsByKey.Clear();
        DefaultStencilCatalog = null;
    }

    private sealed class PopToken : IDisposable
    {
        private VisioDslContext? _context;
        private readonly VisioPage _page;

        internal PopToken(VisioDslContext context, VisioPage page)
        {
            _context = context;
            _page = page;
        }

        public void Dispose()
        {
            _context?.Pop(_page);
            _context = null;
        }
    }
}
