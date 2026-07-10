using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    private static void ReleaseSharedEngineGate(SemaphoreSlim gate, AbandonedOcrOperationTracker operations) {
        IReadOnlyList<Task> pending = operations.GetPendingOperations();
        if (pending.Count == 0) {
            gate.Release();
            return;
        }

        _ = ReleaseSharedEngineGateWhenSettledAsync(gate, pending);
    }

    private static async Task ReleaseSharedEngineGateWhenSettledAsync(SemaphoreSlim gate, IReadOnlyList<Task> pending) {
        try {
            await Task.WhenAll(pending).ConfigureAwait(false);
        } catch {
            // ExecuteCandidateAsync observes provider failures. The shared gate still has to be released.
        } finally {
            gate.Release();
        }
    }

    private sealed class AbandonedOcrOperationTracker {
        private readonly object _sync = new object();
        private readonly List<Task> _operations = new List<Task>();

        internal bool HasPendingOperations {
            get {
                lock (_sync) {
                    return _operations.Any(static operation => !operation.IsCompleted);
                }
            }
        }

        internal void Track(Task? operation) {
            if (operation == null || operation.IsCompleted) return;
            lock (_sync) {
                if (!operation.IsCompleted) _operations.Add(operation);
            }
        }

        internal IReadOnlyList<Task> GetPendingOperations() {
            lock (_sync) {
                return _operations.Where(static operation => !operation.IsCompleted).ToArray();
            }
        }
    }
}
