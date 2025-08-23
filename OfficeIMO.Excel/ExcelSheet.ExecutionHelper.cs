using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel
{
    public partial class ExcelSheet
    {
        /// <summary>
        /// Core execution helper that handles compute/apply split with optional parallelization.
        /// Compute runs without locks; only the apply stage is serialized.
        /// </summary>
        private void ExecuteWithPolicy(
            string opName,
            int itemCount,
            ExecutionMode? overrideMode,
            Action sequentialCore,                // single-threaded path (no locks)
            Action? computeParallel = null,       // parallelizable compute (no DOM)
            Action? applySequential = null,       // serialized DOM apply
            CancellationToken ct = default)
        {
            var policy = EffectiveExecution;
            var mode = overrideMode ?? policy.Mode;
            if (mode == ExecutionMode.Automatic)
                mode = policy.Decide(opName, itemCount);

            if (mode == ExecutionMode.Sequential || computeParallel is null || applySequential is null)
            {
                using (Locking.EnterNoLockScope())
                    sequentialCore();
                return;
            }

            // Parallel: compute without lock
            var po = new ParallelOptions { CancellationToken = ct };
            if (policy.MaxDegreeOfParallelism is int dop && dop > 0)
                po.MaxDegreeOfParallelism = dop;

            computeParallel();

            // Apply once, serialized
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), applySequential);
        }
    }
}