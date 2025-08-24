using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel
{
    public partial class ExcelSheet
    {
        /// <summary>
        /// Core execution helper that handles compute/apply split with optional parallelization.
        /// Compute runs without locks; only the apply stage is serialized under a document-level lock.
        /// </summary>
        /// <param name="opName">A short operation name used for diagnostics and policy decisions.</param>
        /// <param name="itemCount">Approximate number of items to process; used to decide parallel vs sequential.
        /// This does not have to be exact but should reflect relative workload.</param>
        /// <param name="overrideMode">Force a specific execution mode (Sequential/Parallel); null to use policy (Automatic).</param>
        /// <param name="sequentialCore">Single-threaded path. Used either when policy decides sequential or when compute/apply are not provided.</param>
        /// <param name="computeParallel">Compute phase that is safe to run without locks and must not mutate the OpenXML DOM.</param>
        /// <param name="applySequential">Apply phase that writes to the DOM. This runs once under a serialized lock.</param>
        /// <param name="ct">Cancellation token for the compute phase.</param>
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
            // Use existing lock if available; avoid allocating a new lock during disposal edge cases
            System.Threading.ReaderWriterLockSlim? lck = _excelDocument._lock;
            if (lck == null)
            {
                try { lck = _excelDocument.EnsureLock(); } catch { lck = null; }
            }
            Locking.ExecuteWrite(lck, applySequential);
        }
    }
}
