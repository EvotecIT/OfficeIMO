using System;
using System.Collections.Generic;
using System.Management.Automation.Runspaces;

namespace PSWriteOffice.Foundation;

/// <summary>Scopes implicit DSL documents to the current pipeline thread and runspace.</summary>
internal static class OfficeDocumentContext {
    [ThreadStatic] private static Dictionary<Guid, Stack<object>>? _documents;
    private static Guid Key => Runspace.DefaultRunspace?.InstanceId ?? Guid.Empty;
    internal static T? Find<T>() where T : class {
        if (_documents != null && _documents.TryGetValue(Key, out var stack))
            foreach (object value in stack) if (value is T document) return document;
        return null;
    }
    internal static IDisposable Push(object document) {
        Guid key = Key;
        _documents ??= new Dictionary<Guid, Stack<object>>();
        if (!_documents.TryGetValue(key, out var stack)) _documents[key] = stack = new Stack<object>();
        stack.Push(document);
        return new Scope(key, stack);
    }
    private sealed class Scope : IDisposable {
        private readonly Guid _key;
        private Stack<object>? _stack;
        internal Scope(Guid key, Stack<object> stack) { _key = key; _stack = stack; }
        public void Dispose() {
            if (_stack == null) return;
            _stack.Pop();
            if (_stack.Count == 0) _documents!.Remove(_key);
            _stack = null;
        }
    }
}
