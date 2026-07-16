namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Represents the workbook calculation settings stored by BrtCalcProp.</summary>
    internal sealed class XlsbCalculationProperties {
        internal XlsbCalculationProperties(
            uint calculationId,
            uint calculationMode,
            uint iterationCount,
            double iterationDelta,
            int concurrentThreadCount,
            ushort flags) {
            CalculationId = calculationId;
            CalculationMode = calculationMode;
            IterationCount = iterationCount;
            IterationDelta = iterationDelta;
            ConcurrentThreadCount = concurrentThreadCount;
            Flags = flags;
        }

        internal uint CalculationId { get; }

        internal uint CalculationMode { get; }

        internal uint IterationCount { get; }

        internal double IterationDelta { get; }

        internal int ConcurrentThreadCount { get; }

        internal ushort Flags { get; }

        internal bool FullCalculationOnLoad => (Flags & 0x0001) != 0;

        internal bool UsesA1References => (Flags & 0x0002) != 0;

        internal bool IterationEnabled => (Flags & 0x0004) != 0;

        internal bool FullPrecision => (Flags & 0x0008) != 0;

        internal bool CalculationCompleted => (Flags & 0x0010) == 0;

        internal bool CalculationOnSave => (Flags & 0x0020) != 0;

        internal bool ConcurrentCalculation => (Flags & 0x0040) != 0;

        internal bool HasManualConcurrentCount => (Flags & 0x0080) != 0;

        internal bool ForceFullCalculation => (Flags & 0x0100) != 0;
    }
}
