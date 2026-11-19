namespace OfficeIMO.Excel.Xlsb.Biff12;

/// <summary>
/// Converts relative cell references between PtgRef/PtgArea, which store the referenced coordinates, and PtgRefN/PtgAreaN
/// (MS-XLSB 2.5.98.67, 2.5.98.22), which store offsets from an anchor cell. Conditional formatting formulas keep relative
/// references as offsets from the top-left cell of the first range they apply to.
/// </summary>
internal static class XlsbRelativeReferenceTokens {
    private const ushort ColumnRelative = 0x4000;
    private const ushort RowRelative = 0x8000;
    private const int Rows = 1_048_576;
    private const int Columns = 16_384;

    /// <summary>Rewrites relative PtgRef/PtgArea tokens in place as PtgRefN/PtgAreaN offsets from the zero-based anchor.</summary>
    internal static void ToOffsets(byte[] tokens, int anchorRow, int anchorColumn) =>
        XlsbFormulaTokenScanner.Visit(tokens, 0, tokens.Length, (offset, token, operand) => {
            if (token < 0x20) return;
            switch (token & 0x1F) {
                case 0x04 when IsRelative(tokens, operand + 4):
                    tokens[offset] = (byte)(token + 0x08);
                    ToOffset(tokens, operand, operand + 4, anchorRow, anchorColumn);
                    break;
                case 0x05 when IsRelative(tokens, operand + 8) || IsRelative(tokens, operand + 10):
                    tokens[offset] = (byte)(token + 0x08);
                    ToOffset(tokens, operand, operand + 8, anchorRow, anchorColumn);
                    ToOffset(tokens, operand + 4, operand + 10, anchorRow, anchorColumn);
                    break;
            }
        });

    /// <summary>Rewrites PtgRefN/PtgAreaN tokens in place as PtgRef/PtgArea coordinates resolved against the zero-based anchor.</summary>
    internal static void ToCoordinates(byte[] tokens, int anchorRow, int anchorColumn) =>
        XlsbFormulaTokenScanner.Visit(tokens, 0, tokens.Length, (offset, token, operand) => {
            if (token < 0x20) return;
            switch (token & 0x1F) {
                case 0x0C:
                    tokens[offset] = (byte)(token - 0x08);
                    ToCoordinate(tokens, operand, operand + 4, anchorRow, anchorColumn);
                    break;
                case 0x0D:
                    tokens[offset] = (byte)(token - 0x08);
                    ToCoordinate(tokens, operand, operand + 8, anchorRow, anchorColumn);
                    ToCoordinate(tokens, operand + 4, operand + 10, anchorRow, anchorColumn);
                    break;
            }
        });

    /// <summary>True when the tokens hold a relative reference of any kind.</summary>
    internal static bool HasRelativeReferences(byte[] tokens) {
        bool found = false;
        XlsbFormulaTokenScanner.Visit(tokens, 0, tokens.Length, (offset, token, operand) => {
            if (token < 0x20) return;
            switch (token & 0x1F) {
                case 0x04:
                    found |= IsRelative(tokens, operand + 4);
                    break;
                case 0x05:
                    found |= IsRelative(tokens, operand + 8) || IsRelative(tokens, operand + 10);
                    break;
                case 0x0C:
                case 0x0D:
                    found = true;
                    break;
            }
        });
        return found;
    }

    /// <summary>True when the tokens refer to another sheet (PtgRef3d, PtgArea3d and their error forms).</summary>
    internal static bool HasSheetReferences(byte[] tokens) {
        bool found = false;
        XlsbFormulaTokenScanner.Visit(tokens, 0, tokens.Length, (offset, token, operand) => {
            if (token >= 0x20 && (token & 0x1F) is 0x1A or 0x1B or 0x1C or 0x1D) found = true;
        });
        return found;
    }

    private static bool IsRelative(byte[] tokens, int columnOffset) => (ReadUInt16(tokens, columnOffset) & (ColumnRelative | RowRelative)) != 0;

    private static void ToOffset(byte[] tokens, int rowOffset, int columnOffset, int anchorRow, int anchorColumn) {
        ushort column = ReadUInt16(tokens, columnOffset);
        if ((column & RowRelative) != 0) {
            int row = (int)XlsbFormulaTokenScanner.ReadUInt32(tokens, rowOffset);
            XlsbFormulaTokenScanner.WriteUInt32(tokens, rowOffset, unchecked((uint)(row - anchorRow)));
        }
        if ((column & ColumnRelative) != 0) {
            int delta = (column & 0x3FFF) - anchorColumn;
            column = (ushort)((column & (ColumnRelative | RowRelative)) | (delta & 0x3FFF));
            WriteUInt16(tokens, columnOffset, column);
        }
    }

    private static void ToCoordinate(byte[] tokens, int rowOffset, int columnOffset, int anchorRow, int anchorColumn) {
        ushort column = ReadUInt16(tokens, columnOffset);
        if ((column & RowRelative) != 0) {
            int delta = unchecked((int)XlsbFormulaTokenScanner.ReadUInt32(tokens, rowOffset));
            int row = ((anchorRow + delta) % Rows + Rows) % Rows;
            XlsbFormulaTokenScanner.WriteUInt32(tokens, rowOffset, (uint)row);
        }
        if ((column & ColumnRelative) != 0) {
            int delta = column & 0x3FFF;
            if ((delta & 0x2000) != 0) delta -= 0x4000; // ColRelNegative stores a signed 14-bit offset.
            int resolved = ((anchorColumn + delta) % Columns + Columns) % Columns;
            WriteUInt16(tokens, columnOffset, (ushort)((column & (ColumnRelative | RowRelative)) | resolved));
        }
    }

    private static ushort ReadUInt16(byte[] tokens, int offset) => (ushort)(tokens[offset] | tokens[offset + 1] << 8);

    private static void WriteUInt16(byte[] tokens, int offset, ushort value) {
        tokens[offset] = (byte)value;
        tokens[offset + 1] = (byte)(value >> 8);
    }
}
