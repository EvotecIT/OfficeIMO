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
        ref long remainingCalculatorValidationWork,
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
            !program.CanEvaluateDomain(domain, inputCount, outputCount, ref remainingCalculatorValidationWork)) return false;

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
        IReadOnlyList<double>? discontinuities = null;
        if (inputCount == 1) {
            double domainMinimum = Math.Min(domain[0], domain[1]);
            double domainMaximum = Math.Max(domain[0], domain[1]);
            double[] conditionalBoundaries = Array.Empty<double>();
            if (program.HasConditional &&
                !program.TryGetOneInputConditionalBoundaries(domainMinimum, domainMaximum, out conditionalBoundaries)) {
                return false;
            }
            double[] authoredPoints = program.NumericConstants
                .Where(value => value >= domainMinimum && value <= domainMaximum)
                .Concat(program.HasConditional ? conditionalBoundaries : Array.Empty<double>())
                .Distinct()
                .OrderBy(static value => value)
                .ToArray();
            double[] requiredPoints = authoredPoints
                .Concat(domain)
                .Distinct()
                .OrderBy(static value => value)
                .ToArray();
            if (requiredPoints.Length > MaxSuggestedSampleBreakpoints) return false;
            breakpoints = LimitSuggestedPoints(
                CreateUniformBreakpoints(domain, MaxSuggestedSampleBreakpoints).Concat(requiredPoints),
                requiredPoints);
            if (program.HasConditional) {
                discontinuities = conditionalBoundaries
                    .Where(value => value > domainMinimum && value < domainMaximum)
                    .ToArray();
            }
        }

        double[] affineCoefficients = Array.Empty<double>();
        double[] affineConstants = Array.Empty<double>();
        bool hasAffineOutput = inputCount == 1 &&
            program.TryGetOneInputAffineOutput(outputCount, out affineCoefficients, out affineConstants);
        bool affineRangeClippingAbsent = hasAffineOutput &&
            IsAffineRangeClippingAbsent(domain, range, affineCoefficients, affineConstants);
        function = new PdfColorFunction(
            inputCount,
            outputCount,
            domain,
            range,
            (values, output, outputOffset) =>
                program.TryEvaluate(values, output, outputOffset, outputCount),
            breakpoints,
            discontinuities,
            evaluationCost: program.MaximumEvaluationWork,
            requiresAdaptiveShadingSampling: inputCount != 1 || !hasAffineOutput,
            hasUnboundedDiscontinuities: program.HasUnboundedDiscontinuities,
            rangeClippingProvenAbsent: affineRangeClippingAbsent);
        return true;
    }

    private static bool IsAffineRangeClippingAbsent(
        double[] domain,
        double[] range,
        double[] coefficients,
        double[] constants) {
        if (domain.Length != 2 || range.Length < coefficients.Length * 2 ||
            coefficients.Length != constants.Length) return false;
        for (int index = 0; index < coefficients.Length; index++) {
            double atStart = coefficients[index] * domain[0] + constants[index];
            double atEnd = coefficients[index] * domain[1] + constants[index];
            double minimum = Math.Min(atStart, atEnd);
            double maximum = Math.Max(atStart, atEnd);
            if (double.IsNaN(minimum) || double.IsInfinity(minimum) ||
                double.IsNaN(maximum) || double.IsInfinity(maximum) ||
                minimum < range[index * 2] || maximum > range[index * 2 + 1]) return false;
        }
        return true;
    }
}
