namespace OfficeIMO.Odbc {

    internal static class ConnectionSettings {
        internal static string WorkbookPath(string connectionString) {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int index = 0;
            while (index < connectionString.Length) {
                while (index < connectionString.Length && (char.IsWhiteSpace(connectionString[index]) || connectionString[index] == ';')) index++;
                if (index == connectionString.Length) break;
                int equal = connectionString.IndexOf('=', index);
                if (equal < 0) throw new ArgumentException("Connection attributes require key=value syntax.");
                string key = connectionString[index..equal].Trim();
                index = equal + 1;
                while (index < connectionString.Length && char.IsWhiteSpace(connectionString[index])) index++;
                string value;
                if (index < connectionString.Length && connectionString[index] == '{') {
                    index++;
                    System.Text.StringBuilder builder = new System.Text.StringBuilder();
                    bool closed = false;
                    while (index < connectionString.Length) {
                        char character = connectionString[index++];
                        if (character != '}') { builder.Append(character); continue; }
                        if (index < connectionString.Length && connectionString[index] == '}') { builder.Append('}'); index++; continue; }
                        closed = true;
                        break;
                    }
                    if (!closed) throw new ArgumentException("Unterminated connection attribute.");
                    while (index < connectionString.Length && char.IsWhiteSpace(connectionString[index])) index++;
                    if (index < connectionString.Length && connectionString[index] != ';') throw new ArgumentException("Invalid connection attribute separator.");
                    value = builder.ToString();
                } else {
                    int end = connectionString.IndexOf(';', index);
                    if (end < 0) end = connectionString.Length;
                    value = connectionString[index..end].Trim();
                    index = end;
                }
                if (!values.TryAdd(key, value)) throw new ArgumentException("Duplicate connection attribute.");
            }
            foreach (string key in values.Keys) {
                if (key is not null && !new[] { "DRIVER", "DBQ", "READONLY" }.Contains(key, StringComparer.OrdinalIgnoreCase))
                    throw new ArgumentException("Supported connection attributes are DRIVER, DBQ, and ReadOnly.");
            }
            if (values.TryGetValue("ReadOnly", out string? readOnly) && readOnly is not ("1" or "true" or "True"))
                throw new NotSupportedException("Workbook access is read-only.");
            if (!values.TryGetValue("DBQ", out string? path) || string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("DBQ must specify a local workbook path.");
            bool networkOrDevice = path.Length >= 2 && path[0] is '\\' or '/' && path[1] is '\\' or '/';
            if (!System.IO.Path.IsPathFullyQualified(path) || networkOrDevice)
                throw new ArgumentException("DBQ must be an absolute local path; network paths are not supported.");
            if (!System.IO.Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("The ODBC driver accepts .xlsx workbooks.");
            return path;
        }

        internal static void ValidateStoragePath(string path) {
            if (!OperatingSystem.IsWindows()) return;
            string root = System.IO.Path.GetPathRoot(path)!;
            if (new DriveInfo(root).DriveType == DriveType.Network)
                throw new ArgumentException("Mapped network drives are not supported.");
            // Walk from the root so a remote reparse target is rejected before accessing its children.
            string current = root;
            foreach (string component in path[root.Length..].Split(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)) {
                if (component.Length == 0) continue;
                current = System.IO.Path.Combine(current, component);
                if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new ArgumentException("Workbook paths through reparse points are not supported.");
            }
        }
    }
}
