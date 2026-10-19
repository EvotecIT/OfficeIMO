using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    /// <summary>Writes a statistical, run-length-coded VP8L code-length alphabet.</summary>
    private static void WriteVp8lHuffmanTree(ILsbBitWriter writer, byte[] lengths) {
        int symbols = 0, first = -1, second = -1;
        for (int symbol = 0; symbol < lengths.Length; symbol++) {
            if (lengths[symbol] == 0) continue;
            if (first < 0) first = symbol;
            else second = symbol;
            symbols++;
        }
        if (symbols == 1 && first <= 255) {
            WriteSingleSymbolTree(writer, first);
            return;
        }
        if (symbols == 2 && second <= 255) {
            writer.WriteBits(1, 1);
            writer.WriteBits(1, 1);
            writer.WriteBits(first > 1 ? 1U : 0U, 1);
            writer.WriteBits((uint)first, first > 1 ? 8 : 1);
            writer.WriteBits((uint)second, 8);
            return;
        }

        var instructions = new List<Vp8lCodeLengthInstruction>(lengths.Length);
        var frequencies = new int[19];
        for (int index = 0; index < lengths.Length;) {
            int value = lengths[index];
            int run = 1;
            while (index + run < lengths.Length && lengths[index + run] == value) run++;
            index += run;
            if (value != 0) {
                AddVp8lCodeLength(instructions, frequencies, value, 0, 0);
                run--;
            }
            while (run > 0) {
                if (value == 0 && run >= 11) {
                    int repeat = Math.Min(run, 138);
                    AddVp8lCodeLength(instructions, frequencies, 18, 7, repeat - 11);
                    run -= repeat;
                } else if (value == 0 && run >= 3) {
                    int repeat = Math.Min(run, 10);
                    AddVp8lCodeLength(instructions, frequencies, 17, 3, repeat - 3);
                    run -= repeat;
                } else if (value != 0 && run >= 3) {
                    int repeat = Math.Min(run, 6);
                    AddVp8lCodeLength(instructions, frequencies, 16, 2, repeat - 3);
                    run -= repeat;
                } else {
                    AddVp8lCodeLength(instructions, frequencies, value, 0, 0);
                    run--;
                }
            }
        }
        byte[] codeLengths = OfficeHuffmanCodeLengths.Create(frequencies, 7);
        var codes = new Vp8lCodebook(codeLengths);
        int count = Vp8lCodeLengthOrder.Length;
        while (count > 4 && codeLengths[Vp8lCodeLengthOrder[count - 1]] == 0) count--;
        writer.WriteBits(0, 1);
        writer.WriteBits((uint)(count - 4), 4);
        for (int index = 0; index < count; index++) writer.WriteBits(codeLengths[Vp8lCodeLengthOrder[index]], 3);
        writer.WriteBits(0, 1); // RLE expands through the entire destination alphabet
        foreach (Vp8lCodeLengthInstruction instruction in instructions) {
            codes.Write(writer, instruction.Symbol);
            writer.WriteBits((uint)instruction.Value, instruction.Bits);
        }
    }

    private static void AddVp8lCodeLength(List<Vp8lCodeLengthInstruction> instructions,
        int[] frequencies, int symbol, int bits, int value) {
        frequencies[symbol]++;
        instructions.Add(new Vp8lCodeLengthInstruction(symbol, bits, value));
    }

    private readonly struct Vp8lCodeLengthInstruction {
        internal Vp8lCodeLengthInstruction(int symbol, int bits, int value) {
            Symbol = symbol;
            Bits = bits;
            Value = value;
        }
        internal int Symbol { get; }
        internal int Bits { get; }
        internal int Value { get; }
    }
}
