using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

internal sealed class WordDslContext : IDisposable
{
    private static readonly AsyncLocal<WordDslContext?> CurrentScope = new();
    private readonly Stack<object> _scopes = new();
    private readonly Dictionary<WordList, WordParagraph?> _listAnchors = new();
    private readonly Dictionary<WordTable, IReadOnlyList<object>> _tableSources = new();
    private readonly Dictionary<WordTable, List<WordTableConditionModel>> _tableConditions = new();
    private bool _initialSectionConsumed;

    private WordDslContext(WordDocument document)
    {
        Document = document ?? throw new ArgumentNullException(nameof(document));
    }

    public WordDocument Document { get; }

    public static WordDslContext? Current => CurrentScope.Value;

    public static WordDslContext Enter(WordDocument document)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        if (CurrentScope.Value != null)
        {
            throw new InvalidOperationException("A Word DSL scope is already active on this runspace.");
        }

        var scope = new WordDslContext(document);
        CurrentScope.Value = scope;
        return scope;
    }

    public static WordDslContext Require(PSCmdlet caller)
    {
        var scope = CurrentScope.Value;
        if (scope == null)
        {
            throw new InvalidOperationException(
                $"'{caller.MyInvocation.InvocationName}' must run inside a Word DSL script block.");
        }

        return scope;
    }

    public void Dispose()
    {
        if (CurrentScope.Value == this)
        {
            CurrentScope.Value = null;
        }
        _scopes.Clear();
    }

    public IDisposable Push(object scope)
    {
        if (scope == null)
        {
            throw new ArgumentNullException(nameof(scope));
        }

        _scopes.Push(scope);
        return new PopToken(this, scope);
    }

    private void Pop(object scope)
    {
        if (_scopes.Count == 0)
        {
            return;
        }

        if (ReferenceEquals(_scopes.Peek(), scope))
        {
            _scopes.Pop();
        }
    }

    private sealed class PopToken : IDisposable
    {
        private WordDslContext? _context;
        private readonly object _scope;

        public PopToken(WordDslContext context, object scope)
        {
            _context = context;
            _scope = scope;
        }

        public void Dispose()
        {
            _context?.Pop(_scope);
            _context = null;
        }
    }

    public WordSection? CurrentSection => _scopes.OfType<WordSection>().FirstOrDefault();
    public WordHeader? CurrentHeader => _scopes.OfType<WordHeader>().FirstOrDefault();
    public WordFooter? CurrentFooter => _scopes.OfType<WordFooter>().FirstOrDefault();
    public WordParagraph? CurrentParagraph => _scopes.OfType<WordParagraph>().FirstOrDefault();
    public WordTable? CurrentTable => _scopes.OfType<WordTable>().FirstOrDefault();
    public WordTableCell? CurrentTableCell => _scopes.OfType<WordTableCell>().FirstOrDefault();
    public WordList? CurrentList => _scopes.OfType<WordList>().FirstOrDefault();

    public WordSection AcquireSection(WordSectionBreakType? breakType = null)
    {
        if (!_initialSectionConsumed && Document.Sections.Count > 0)
        {
            _initialSectionConsumed = true;
            return Document.Sections.Last();
        }

        _initialSectionConsumed = true;
        return breakType.HasValue
            ? Document.AddSection(breakType.Value)
            : Document.AddSection();
    }

    public WordSection RequireSection()
    {
        return CurrentSection ?? RequireImplicitSection();
    }

    private WordSection RequireImplicitSection()
    {
        _initialSectionConsumed = true;
        return Document.Sections.Count > 0
            ? Document.Sections.Last()
            : Document.AddSection();
    }

    public object RequireParagraphHost()
    {
        if (CurrentTableCell != null)
        {
            return CurrentTableCell;
        }
        if (CurrentHeader != null)
        {
            return CurrentHeader;
        }
        if (CurrentFooter != null)
        {
            return CurrentFooter;
        }

        return RequireSection();
    }

    public WordTable? ResolveCurrentTable()
    {
        return CurrentTable;
    }

    public void RegisterListAnchor(WordList list, WordParagraph? anchor)
    {
        if (list == null)
        {
            throw new ArgumentNullException(nameof(list));
        }
        _listAnchors[list] = anchor;
    }

    public WordParagraph? ConsumeListAnchor(WordList list)
    {
        if (list == null)
        {
            throw new ArgumentNullException(nameof(list));
        }
        if (_listAnchors.TryGetValue(list, out var anchor))
        {
            _listAnchors.Remove(list);
            return anchor;
        }
        return null;
    }

    public void ClearListAnchor(WordList list)
    {
        if (list == null)
        {
            throw new ArgumentNullException(nameof(list));
        }
        _listAnchors.Remove(list);
    }

    public void RegisterTableSource(WordTable table, IReadOnlyList<object> rows)
    {
        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }
        _tableSources[table] = rows;
    }

    public IReadOnlyList<object> GetTableSource(WordTable table)
    {
        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }
        return _tableSources.TryGetValue(table, out var rows)
            ? rows
            : Array.Empty<object>();
    }

    public void ClearTableSource(WordTable table)
    {
        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }
        _tableSources.Remove(table);
    }

    public void AddTableCondition(WordTable table, WordTableConditionModel model)
    {
        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }
        if (!_tableConditions.TryGetValue(table, out var list))
        {
            list = new List<WordTableConditionModel>();
            _tableConditions[table] = list;
        }
        list.Add(model);
    }

    public IReadOnlyList<WordTableConditionModel> ConsumeTableConditions(WordTable table)
    {
        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }
        if (_tableConditions.TryGetValue(table, out var list))
        {
            _tableConditions.Remove(table);
            return list;
        }
        return Array.Empty<WordTableConditionModel>();
    }
}
