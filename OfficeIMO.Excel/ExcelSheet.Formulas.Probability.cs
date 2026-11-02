namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // The two vectors share the existing range budget. Only numeric vectors are
        // evaluated here; mixed-type coercion remains outside the qualified subset.
        private bool TryEvaluateProbabilityValue(string args, out FormulaArgumentValue result) {
            result = default;
            IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (tokens.Count < 3 || tokens.Count > 4 || tokens.Any(string.IsNullOrWhiteSpace)) return false;
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            if (!TryResolveFormulaRange(tokens[0], out List<FormulaArgumentValue> outcomes, ref remainingCellBudget)
                || !TryResolveFormulaRange(tokens[1], out List<FormulaArgumentValue> probabilities, ref remainingCellBudget)) return false;
            if (!TryResolveFormulaValueArgument(tokens[2], out FormulaArgumentValue lower)) return false;
            FormulaArgumentValue upper = lower;
            if (tokens.Count == 4 && !TryResolveFormulaValueArgument(tokens[3], out upper)) return false;
            if ((!lower.IsError && !IsFiniteFormulaNumber(lower)) || (!upper.IsError && !IsFiniteFormulaNumber(upper))) return false;
            if (outcomes.Concat(probabilities).Any(value => value.IsUnresolvedFormula
                || (!value.IsError && (!value.IsNumericAggregateValue || !IsFiniteFormulaNumber(value))))) return false;
            if (_formulaReferenceError.IsError) { result = _formulaReferenceError; return true; }
            if (lower.IsError) { result = lower; return true; }
            if (upper.IsError) { result = upper; return true; }
            if (outcomes.Count != probabilities.Count) { result = FormulaArgumentValue.Error("#N/A"); return true; }
            double total = 0d, selected = 0d;
            for (int index = 0; index < outcomes.Count; index++) {
                FormulaArgumentValue outcome = outcomes[index], probability = probabilities[index];
                if (outcome.IsUnresolvedFormula || probability.IsUnresolvedFormula) return false;
                if (outcome.IsError) { result = outcome; return true; }
                if (probability.IsError) { result = probability; return true; }
                if (!outcome.IsNumericAggregateValue || !probability.IsNumericAggregateValue
                    || !IsFiniteFormulaNumber(outcome) || !IsFiniteFormulaNumber(probability)) return false;
                double p = probability.Number!.Value;
                if (p < 0d || p > 1d) { result = FormulaArgumentValue.Error("#NUM!"); return true; }
                total += p;
                if (outcome.Number!.Value >= lower.Number!.Value && outcome.Number.Value <= upper.Number!.Value) selected += p;
            }
            // Permit ordinary floating-point addition noise, never normalize an invalid distribution.
            if (Math.Abs(total - 1d) > 1e-12) { result = FormulaArgumentValue.Error("#NUM!"); return true; }
            result = new FormulaArgumentValue(selected, InvariantNumberText.Get(selected));
            return true;
        }

        private static bool IsFiniteFormulaNumber(FormulaArgumentValue value) => value.Number.HasValue
            && !double.IsNaN(value.Number.Value) && !double.IsInfinity(value.Number.Value);

    }
}
