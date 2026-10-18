using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Services.OpenDocument;

internal sealed class OpenDocumentDslContext : IDisposable {
    private static readonly AsyncLocal<OpenDocumentDslContext?> CurrentScope = new();
    private readonly Stack<object> _scopes = new();

    private OpenDocumentDslContext(OdfDocument document) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
    }

    internal OdfDocument Document { get; }
    internal static OpenDocumentDslContext? Current => CurrentScope.Value;

    internal static OpenDocumentDslContext Enter(OdfDocument document) {
        if (CurrentScope.Value != null) {
            throw new InvalidOperationException("An OpenDocument DSL scope is already active on this runspace.");
        }

        var context = new OpenDocumentDslContext(document);
        CurrentScope.Value = context;
        return context;
    }

    internal static OpenDocumentDslContext Require(PSCmdlet caller) => CurrentScope.Value
        ?? throw new PSInvalidOperationException(
            $"'{caller.MyInvocation.InvocationName}' requires -Document or an active New-OfficeOpenDocument -Content scope.");

    internal T RequireDocument<T>(PSCmdlet caller, string kindName) where T : OdfDocument => Document as T
        ?? throw new PSInvalidOperationException(
            $"'{caller.MyInvocation.InvocationName}' requires an OpenDocument {kindName} document.");

    internal OdsSheet RequireSheet() => _scopes.OfType<OdsSheet>().FirstOrDefault()
        ?? throw new PSInvalidOperationException("No OpenDocument worksheet context is active. Use Add-OfficeOpenDocumentSheet -Content first or pass -Sheet.");

    internal OdpSlide RequireSlide() => _scopes.OfType<OdpSlide>().FirstOrDefault()
        ?? throw new PSInvalidOperationException("No OpenDocument slide context is active. Use Add-OfficeOpenDocumentSlide -Content first or pass -Slide.");

    internal IDisposable Push(object scope) {
        _scopes.Push(scope ?? throw new ArgumentNullException(nameof(scope)));
        return new PopToken(this, scope);
    }

    public void Dispose() {
        if (CurrentScope.Value == this) CurrentScope.Value = null;
        _scopes.Clear();
    }

    private void Pop(object scope) {
        if (_scopes.Count > 0 && ReferenceEquals(_scopes.Peek(), scope)) _scopes.Pop();
    }

    private sealed class PopToken : IDisposable {
        private OpenDocumentDslContext? _context;
        private readonly object _scope;

        internal PopToken(OpenDocumentDslContext context, object scope) {
            _context = context;
            _scope = scope;
        }

        public void Dispose() {
            _context?.Pop(_scope);
            _context = null;
        }
    }
}
