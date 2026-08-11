using System.IO;

namespace OfficeIMO.Pdf;

internal static partial class PdfColorSpaceFunctionResolver {
    private static bool TryCreateCalculatorFunction(
        PdfStream? stream,
        PdfDictionary dictionary,
        int inputCount,
        int outputCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        ref long retainedFunctionBytes,
        out PdfColorFunction function) {
        function = null!;
        if (stream == null ||
            inputCount > 256 ||
            outputCount > 256 ||
            !TryReadIntervals(dictionary, "Domain", inputCount, objects, allowEqual: true, required: true, out double[] domain) ||
            !TryReadIntervals(dictionary, "Range", outputCount, objects, allowEqual: true, required: true, out double[] range)) return false;

        long remainingBytes = maxDecodedStreamBytes - retainedFunctionBytes;
        if (remainingBytes <= 0L) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maxDecodedStreamBytes, retainedFunctionBytes + 1L);
        }

        int decodeLimit = (int)Math.Min(
            Math.Min(remainingBytes, int.MaxValue),
            PdfCalculatorProgram.MaxProgramBytes + 1L);
        byte[] bytes;
        try {
            bytes = Filters.StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, decodeLimit);
        } catch (InvalidDataException) {
            return false;
        }
        if (bytes.Length > PdfCalculatorProgram.MaxProgramBytes ||
            !PdfCalculatorProgram.TryParse(bytes, out PdfCalculatorProgram program) ||
            !program.CanEvaluateDomain(domain, inputCount, outputCount)) return false;

        long totalRetainedBytes;
        try {
            totalRetainedBytes = checked(retainedFunctionBytes + program.RetainedBytes);
        } catch (OverflowException) {
            return false;
        }
        if (totalRetainedBytes > maxDecodedStreamBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maxDecodedStreamBytes, totalRetainedBytes);
        }
        retainedFunctionBytes = totalRetainedBytes;

        IReadOnlyList<double>? breakpoints = null;
        if (inputCount == 1) {
            IEnumerable<double> authoredPoints = program.NumericConstants
                .Where(value => value >= domain[0] && value <= domain[1]);
            breakpoints = LimitSuggestedPoints(
                CreateUniformBreakpoints(domain, MaxSuggestedSampleBreakpoints).Concat(authoredPoints).Concat(domain),
                domain);
        }

        function = new PdfColorFunction(
            inputCount,
            outputCount,
            domain,
            range,
            values => program.Evaluate(values, outputCount),
            breakpoints,
            evaluationCost: program.MaximumEvaluationWork);
        return true;
    }
}
