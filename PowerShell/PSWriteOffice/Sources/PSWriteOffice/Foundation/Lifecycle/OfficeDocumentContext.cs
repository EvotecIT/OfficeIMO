using System;
using System.Collections.Generic;
using System.Management.Automation.Runspaces;

namespace PSWriteOffice.Foundation;

/// <summary>Scopes implicit DSL documents to the current pipeline thread and runspace.</summary>
internal static class OfficeDocumentContext {
    [ThreadStatic] private static Dictionary<Guid, Stack<Scope>>? _documents;
    private static Guid Key => Runspace.DefaultRunspace?.InstanceId ?? Guid.Empty;
    internal static T? Find<T>() where T : class {
        if (_documents != null && _documents.TryGetValue(Key, out var stack))
            foreach (Scope scope in stack) if (scope.Document is T document) return document;
        return null;
    }
    internal static Scope Push(object document) {
        Guid key = Key;
        _documents ??= new Dictionary<Guid, Stack<Scope>>();
        if (!_documents.TryGetValue(key, out var stack)) _documents[key] = stack = new Stack<Scope>();
        var scope = new Scope(key, stack, document);
        stack.Push(scope);
        return scope;
    }
    /// <summary>Adopts a native immutable edit only in the scope that owns its input.</summary>
    internal static bool Replace<T>(T original, T replacement) where T : class, IDisposable {
        if (_documents == null || !_documents.TryGetValue(Key, out var stack)) return false;
        foreach (Scope scope in stack) {
            if (!ReferenceEquals(scope.Document, original)) continue;
            scope.Replace(replacement);
            return true;
        }
        return false;
    }
    internal sealed class Scope : IDisposable {
        private readonly Guid _key;
        private readonly object _original;
        private Stack<Scope>? _stack;
        internal object Document { get; private set; }
        internal Scope(Guid key, Stack<Scope> stack, object document) {
            _key = key;
            _stack = stack;
            _original = Document = document;
        }
        internal void Replace(object replacement) {
            object previous = Document;
            Document = replacement;
            if (!ReferenceEquals(previous, _original) && !ReferenceEquals(previous, replacement))
                (previous as IDisposable)?.Dispose();
        }
        public void Dispose() {
            if (_stack == null) return;
            try {
                if (!ReferenceEquals(Document, _original)) (Document as IDisposable)?.Dispose();
            } finally {
                _stack.Pop();
                if (_stack.Count == 0) _documents!.Remove(_key);
                _stack = null;
            }
        }
    }
}
