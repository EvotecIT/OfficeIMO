namespace OfficeIMO.Access {
    /// <summary>Generation-specific table coordinates remain private behind the shared public table model.</summary>
    internal sealed class AccessNativeTable {
        internal AccessNativeTable(AccessNativeDatabase database, int page, string name) { Database = database; DefinitionPage = page; Name = name; }
        internal readonly AccessNativeDatabase Database;
        internal readonly int DefinitionPage;
        internal readonly string Name;
        internal long RowCount;
        internal int MaxColumns, MaxVariableColumns;
        internal uint OwnedPages;
        internal readonly List<AccessNativeColumn> Columns = new List<AccessNativeColumn>();
        internal readonly List<AccessNativeIndex> Indexes = new List<AccessNativeIndex>();
        internal AccessTable? Model;
    }

    internal sealed class AccessNativeColumn {
        internal AccessComplexDefinition? ComplexDefinition;
        internal bool RedactConnection;
        internal string Name = string.Empty;
        internal byte Type, Flags, ExtraFlags, Precision, Scale;
        internal int Number, VariableIndex, FixedOffset, Size, ComplexId;
        internal bool Variable => (Flags & 1) == 0;
        internal bool Calculated => (ExtraFlags & 0xc0) != 0;
        internal AccessColumn? Model;
    }

    internal sealed class AccessNativeIndex {
        internal string Name = string.Empty;
        internal byte Type, Flags;
        internal int RootPage, Number, RelatedTable, RelatedIndex;
        internal bool CascadeUpdates, CascadeDeletes;
        internal AccessNativeColumn[] Columns = Array.Empty<AccessNativeColumn>();
        internal bool[] Descending = Array.Empty<bool>();
    }
}
