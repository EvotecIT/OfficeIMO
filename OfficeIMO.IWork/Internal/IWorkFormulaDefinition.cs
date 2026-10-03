namespace OfficeIMO.IWork.Internal;

/// <summary>Retains an immutable, bounded source expression for identity-based destination rendering.</summary>
internal sealed class IWorkFormulaDefinition {
    private readonly IWorkWireMessage _formula;
    private readonly int _row, _column, _maximumNodes, _maximumCharacters;
    internal IWorkFormulaDefinition(IWorkWireMessage formula, int row, int column, int maximumNodes, int maximumCharacters) {
        _formula = formula; _row = row; _column = column;
        _maximumNodes = maximumNodes; _maximumCharacters = maximumCharacters;
    }
    internal IWorkFormulaResult Render(IReadOnlyDictionary<Guid, IWorkFormulaTableBinding> qualifiers, IWorkProjectionBudget budget, IWorkFormulaTableBinding? owningTable = null) {
        budget.AddFormulaRenderingOperations(IWorkFormulaReader.MeasureRenderingOperations(_formula, _maximumNodes));
        return IWorkFormulaReader.Render(_formula, _row, _column, _maximumNodes, _maximumCharacters, qualifiers, owningTable);
    }
}
