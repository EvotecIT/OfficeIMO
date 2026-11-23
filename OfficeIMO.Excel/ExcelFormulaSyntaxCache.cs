namespace OfficeIMO.Excel {
    /// <summary>
    /// Formula syntax a workbook's calculation passes reuse: function calls, argument lists and compiled scalar
    /// expressions, each keyed by the exact text it was parsed from. Every entry is determined by its text alone, so
    /// an edited formula, being new text, is parsed afresh and an old entry is never wrong, only unused.
    /// A pass installs the cache for its thread while it holds the workbook's write lock; evaluation outside a pass
    /// parses as before. The cache is cleared once it holds too much text, so a long-lived workbook stays bounded.
    /// </summary>
    internal sealed class ExcelFormulaSyntaxCache {
        private const int MaximumEntries = 131072;
        private const long MaximumKeyCharacters = 8L * 1024 * 1024;

        [ThreadStatic]
        private static ExcelFormulaSyntaxCache? _current;

        private readonly Dictionary<string, ExcelFormulaFunctionCallSyntax?> _calls = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string[]> _arguments = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string[]> _argumentsWithEmpty = new(StringComparer.Ordinal);
        private readonly Dictionary<string, object?> _scalars = new(StringComparer.Ordinal);
        private int _entries;
        private long _keyCharacters;

        /// <summary>The cache installed for the calling thread's calculation pass, if one is running.</summary>
        internal static ExcelFormulaSyntaxCache? Current => _current;

        /// <summary>Installs this cache for the calling thread until the returned scope is disposed.</summary>
        internal Scope Enter() {
            ExcelFormulaSyntaxCache? previous = _current;
            _current = this;
            return new Scope(previous);
        }

        internal bool TryGetCall(string text, out ExcelFormulaFunctionCallSyntax? call) => _calls.TryGetValue(text, out call);

        internal void AddCall(string text, ExcelFormulaFunctionCallSyntax? call) {
            Reserve(text);
            _calls[text] = call;
        }

        internal bool TryGetArguments(string text, bool preserveEmpty, out string[] arguments) =>
            (preserveEmpty ? _argumentsWithEmpty : _arguments).TryGetValue(text, out arguments!);

        internal void AddArguments(string text, bool preserveEmpty, string[] arguments) {
            Reserve(text);
            (preserveEmpty ? _argumentsWithEmpty : _arguments)[text] = arguments;
        }

        /// <summary>Compiled scalar expressions; a null entry records text the compiler does not accept.</summary>
        internal bool TryGetScalar(string text, out object? compiled) => _scalars.TryGetValue(text, out compiled);

        internal void AddScalar(string text, object? compiled) {
            Reserve(text);
            _scalars[text] = compiled;
        }

        private void Reserve(string text) {
            if (_entries >= MaximumEntries || _keyCharacters + text.Length > MaximumKeyCharacters) {
                _calls.Clear();
                _arguments.Clear();
                _argumentsWithEmpty.Clear();
                _scalars.Clear();
                _entries = 0;
                _keyCharacters = 0;
            }
            _entries++;
            _keyCharacters += text.Length;
        }

        internal readonly struct Scope : IDisposable {
            private readonly ExcelFormulaSyntaxCache? _previous;

            internal Scope(ExcelFormulaSyntaxCache? previous) {
                _previous = previous;
            }

            public void Dispose() {
                _current = _previous;
            }
        }
    }

    public partial class ExcelDocument {
        private ExcelFormulaSyntaxCache? _formulaSyntaxCache;

        /// <summary>Formula syntax kept across this workbook's calculation passes; use it only under the write lock.</summary>
        internal ExcelFormulaSyntaxCache FormulaSyntaxCache => _formulaSyntaxCache ??= new ExcelFormulaSyntaxCache();
    }
}
