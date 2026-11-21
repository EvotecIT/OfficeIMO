namespace OfficeIMO.Excel {
    /// <summary>
    /// Bounds background read helpers across every concurrently opening reader in the process.
    /// Readers that cannot reserve a helper index and decode on their own thread instead of
    /// oversubscribing a busy host.
    /// </summary>
    internal static class ExcelParallelReadBudget {
        private static int s_active;

        /// <summary>Half the processors this process may use, less one for the reader's own work.</summary>
        private static int Limit => Math.Max(0, Environment.ProcessorCount / 2 - 1);

        /// <summary>Reserves up to <paramref name="requested"/> helpers and returns how many were granted.</summary>
        internal static int Reserve(int requested) {
            while (requested > 0) {
                int active = Volatile.Read(ref s_active);
                int granted = Math.Min(requested, Limit - active);
                if (granted <= 0) return 0;
                if (Interlocked.CompareExchange(ref s_active, active + granted, active) == active) return granted;
            }
            return 0;
        }

        internal static void Release(int count) {
            if (count > 0) Interlocked.Add(ref s_active, -count);
        }
    }
}
