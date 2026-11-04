using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Data {
    public partial class ObjectFlattenerOptions {
        // Native table hosts share display headers without changing the paths used to select values.
        internal OfficeTableOptions? TableOptions { get; set; }

        internal HashSet<string>? ExplicitColumnLookup { get; private set; }

        internal HashSet<string>? ExplicitColumnAncestorLookup { get; private set; }

        internal ObjectFlattenerOptions CreateProjectionSnapshot() {
            var snapshot = (ObjectFlattenerOptions)MemberwiseClone();
            snapshot.Formatters = new Dictionary<string, Func<object?, object?>>(Formatters, StringComparer.OrdinalIgnoreCase);
            snapshot.Columns = Columns?.ToArray();
            if (snapshot.Columns == null || snapshot.Columns.Length == 0) {
                snapshot.ExplicitColumnLookup = null;
                snapshot.ExplicitColumnAncestorLookup = null;
                return snapshot;
            }

            snapshot.ExplicitColumnLookup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            snapshot.ExplicitColumnAncestorLookup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string column in snapshot.Columns) {
                if (string.IsNullOrWhiteSpace(column)) continue;
                snapshot.ExplicitColumnLookup.Add(column);
                int separator = column.IndexOf('.');
                while (separator > 0) {
                    snapshot.ExplicitColumnAncestorLookup.Add(column.Substring(0, separator));
                    separator = column.IndexOf('.', separator + 1);
                }
            }
            return snapshot;
        }
    }
}
