namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Describes why a BIFF formula token stream could not be decoded into Open XML formula text.
    /// </summary>
    internal sealed class BiffFormulaReadFailure {
        private BiffFormulaReadFailure(string detailCode, string description, byte? token = null, int? tokenOffset = null) {
            DetailCode = detailCode;
            TokenName = token.HasValue ? BiffFormulaTokenInfo.GetTokenName(token.Value) : null;
            Description = AppendTokenLocation(description, token, TokenName, tokenOffset);
            Token = token;
            TokenOffset = tokenOffset;
        }

        /// <summary>
        /// Gets a stable detail key suitable for import report grouping.
        /// </summary>
        internal string DetailCode { get; }

        /// <summary>
        /// Gets a short human-readable description of the unsupported formula shape.
        /// </summary>
        internal string Description { get; }

        /// <summary>
        /// Gets the BIFF formula token byte that exposed the failure, when the token stream was readable that far.
        /// </summary>
        internal byte? Token { get; }

        /// <summary>
        /// Gets the stable BIFF parsed-formula token name, when a token byte is available.
        /// </summary>
        internal string? TokenName { get; }

        /// <summary>
        /// Gets the zero-based offset of the token within the parsed-expression token stream, when available.
        /// </summary>
        internal int? TokenOffset { get; }

        internal static BiffFormulaReadFailure InvalidPayload(string reason) =>
            new("FormulaInvalidPayload", reason);

        internal static BiffFormulaReadFailure InvalidPayload(string reason, byte token, int tokenOffset) =>
            new("FormulaInvalidPayload", reason, token, tokenOffset);

        internal static BiffFormulaReadFailure Stack(string reason) =>
            new("FormulaStackShape", reason);

        internal static BiffFormulaReadFailure Stack(string reason, byte token, int tokenOffset) =>
            new("FormulaStackShape", reason, token, tokenOffset);

        internal static BiffFormulaReadFailure UnsupportedToken(byte token, int tokenOffset) =>
            new($"FormulaToken0x{token:X2}", $"Unsupported formula token {BiffFormulaTokenInfo.GetTokenName(token)} (0x{token:X2}).", token, tokenOffset);

        internal static BiffFormulaReadFailure UnsupportedAttribute(byte attribute, byte token, int tokenOffset) =>
            new($"FormulaAttribute0x{attribute:X2}", $"Unsupported formula attribute 0x{attribute:X2}.", token, tokenOffset);

        internal static BiffFormulaReadFailure UnsupportedFixedFunction(ushort functionId, byte token, int tokenOffset) =>
            new($"FormulaFixedFunction0x{functionId:X4}", $"Unsupported fixed-arity formula function id 0x{functionId:X4}.", token, tokenOffset);

        internal static BiffFormulaReadFailure UnsupportedVariableFunction(ushort functionId, bool isCetabFunction, byte token, int tokenOffset) =>
            isCetabFunction
                ? new($"FormulaCetabFunction0x{functionId:X4}", $"Unsupported add-in formula function id 0x{functionId:X4}.", token, tokenOffset)
                : new($"FormulaVariableFunction0x{functionId:X4}", $"Unsupported variable-arity formula function id 0x{functionId:X4}.", token, tokenOffset);

        internal static BiffFormulaReadFailure UnsupportedFunctionArguments(ushort functionId, int parameterCount, byte token, int tokenOffset) =>
            new($"FormulaFunction0x{functionId:X4}Args{parameterCount}", $"Unsupported argument count {parameterCount} for formula function id 0x{functionId:X4}.", token, tokenOffset);

        internal static BiffFormulaReadFailure FunctionStackUnderflow(ushort functionId, string functionName, int parameterCount, int availableCount, byte token, int tokenOffset) =>
            new($"FormulaFunction0x{functionId:X4}StackUnderflow", $"Formula function {functionName} (0x{functionId:X4}) expected {parameterCount} stack operands but only {availableCount} were available.", token, tokenOffset);

        internal static BiffFormulaReadFailure DefinedName(uint oneBasedNameIndex, byte token, int tokenOffset) =>
            new($"FormulaDefinedName{oneBasedNameIndex}", $"Formula defined-name operand {oneBasedNameIndex} could not be resolved.", token, tokenOffset);

        internal static BiffFormulaReadFailure ExternalName(ushort externSheetIndex, uint oneBasedNameIndex, byte token, int tokenOffset) =>
            new("FormulaExternalName", $"Formula external defined-name operand {oneBasedNameIndex} through ExternSheet index {externSheetIndex} could not be resolved.", token, tokenOffset);

        internal static BiffFormulaReadFailure Reference(string detailCode, string reason, byte token, int tokenOffset) =>
            new(detailCode, reason, token, tokenOffset);

        private static string AppendTokenLocation(string description, byte? token, string? tokenName, int? tokenOffset) {
            if (!token.HasValue || !tokenOffset.HasValue) {
                return description;
            }

            string tokenDisplay = string.IsNullOrWhiteSpace(tokenName)
                ? $"Token 0x{token.Value:X2}"
                : $"Token {tokenName} (0x{token.Value:X2})";
            return $"{description.TrimEnd()} {tokenDisplay} at parsed-expression offset {tokenOffset.Value}.";
        }
    }
}
