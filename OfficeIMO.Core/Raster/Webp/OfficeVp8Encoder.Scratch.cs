// Scratch belongs to one encoding operation and never escapes the encoder.
namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private readonly byte[] _macroblocksPredictedMacroblock = new byte[256];
    private readonly byte[] _macroblocksYBackup = new byte[MacroblockSize * MacroblockSize];
    private readonly byte[] _macroblocksUBackup = new byte[(MacroblockSize / 2) * (MacroblockSize / 2)];
    private readonly byte[] _macroblocksVBackup = new byte[(MacroblockSize / 2) * (MacroblockSize / 2)];
    private readonly byte[] _blocksMacroPrediction = new byte[256];
    private readonly double[] _blocksDcValues = new double[MacroblockSubBlockCount];
    private readonly int[] _blocksYQuant = new int[MacroblockSubBlockCount * CoefficientsPerBlock];
    private readonly byte[] _blocksPredicted = new byte[BlockSize * BlockSize];
    private readonly int[] _blocksResidual = new int[CoefficientsPerBlock];
    private readonly double[] _blocksCoeffs = new double[CoefficientsPerBlock];
    private readonly int[] _blocksDequantCoeffs = new int[CoefficientsPerBlock];
    private readonly double[] _blocksY2Coeff = new double[CoefficientsPerBlock];
    private readonly int[] _blocksY2Quant = new int[CoefficientsPerBlock];
    private readonly int[] _blocksCoeffTokens = new int[CoefficientsPerBlock];
    private readonly int[] _blocksY2Dequant = new int[CoefficientsPerBlock];
    private readonly double[] _blocksDequantized = new double[CoefficientsPerBlock];
    private readonly int[] _blocksCoefficients = new int[CoefficientsPerBlock];
    private readonly byte[] _predictionModesPredicted = new byte[BlockSize * BlockSize];
    private readonly byte[] _predictionModesPredictedU = new byte[64];
    private readonly byte[] _predictionModesPredictedV = new byte[64];
    private readonly byte[] _predictionPredicted = new byte[256];
    private readonly int[] _transformTemporaryInt = new int[16];
    private readonly double[] _transformTemporaryDouble = new double[16];
    private readonly int[] _macroblockModes = new int[MacroblockSubBlockCount];
}
