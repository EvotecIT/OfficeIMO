namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Describes why a BIFF formula token stream could not be decoded into Open XML formula text.
    /// </summary>
    internal sealed class BiffFormulaReadFailure {
        private BiffFormulaReadFailure(string detailCode, string description) {
            DetailCode = detailCode;
            Description = description;
        }

        /// <summary>
        /// Gets a stable detail key suitable for import report grouping.
        /// </summary>
        internal string DetailCode { get; }

        /// <summary>
        /// Gets a short human-readable description of the unsupported formula shape.
        /// </summary>
        internal string Description { get; }

        internal static BiffFormulaReadFailure InvalidPayload(string reason) =>
            new("FormulaInvalidPayload", reason);

        internal static BiffFormulaReadFailure Stack(string reason) =>
            new("FormulaStackShape", reason);

        internal static BiffFormulaReadFailure UnsupportedToken(byte token) =>
            new($"FormulaToken0x{token:X2}", $"Unsupported formula token 0x{token:X2}.");

        internal static BiffFormulaReadFailure UnsupportedAttribute(byte attribute) =>
            new($"FormulaAttribute0x{attribute:X2}", $"Unsupported formula attribute 0x{attribute:X2}.");

        internal static BiffFormulaReadFailure UnsupportedFixedFunction(ushort functionId) =>
            new($"FormulaFixedFunction0x{functionId:X4}", $"Unsupported fixed-arity formula function id 0x{functionId:X4}.");

        internal static BiffFormulaReadFailure UnsupportedVariableFunction(ushort functionId, bool isCetabFunction) =>
            isCetabFunction
                ? new($"FormulaCetabFunction0x{functionId:X4}", $"Unsupported add-in formula function id 0x{functionId:X4}.")
                : new($"FormulaVariableFunction0x{functionId:X4}", $"Unsupported variable-arity formula function id 0x{functionId:X4}.");

        internal static BiffFormulaReadFailure UnsupportedFunctionArguments(ushort functionId, int parameterCount) =>
            new($"FormulaFunction0x{functionId:X4}Args{parameterCount}", $"Unsupported argument count {parameterCount} for formula function id 0x{functionId:X4}.");

        internal static BiffFormulaReadFailure DefinedName(uint oneBasedNameIndex) =>
            new($"FormulaDefinedName{oneBasedNameIndex}", $"Formula defined-name operand {oneBasedNameIndex} could not be resolved.");

        internal static BiffFormulaReadFailure Reference(string reason) =>
            new("FormulaReference", reason);
    }
}
