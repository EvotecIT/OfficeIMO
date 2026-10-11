using System.Data.Odbc;
using System.Runtime.InteropServices;
using Microsoft.Win32;

if (args.Length != 2) throw new ArgumentException("Expected connection string and private verification registry hive.");
using RegistryKey hive = Registry.CurrentUser.OpenSubKey(args[1]) ?? throw new InvalidOperationException("Missing verification hive.");
if (Native.RegOverridePredefKey(unchecked((nint)(long)0xffffffff80000002), hive.Handle.DangerousGetHandle()) != 0)
    throw new InvalidOperationException("Cannot redirect the verification process's registry.");
try {
    using var connection = new OdbcConnection(args[0]);
    connection.Open();
    using var command = connection.CreateCommand();
    command.CommandText = "SELECT \"Name\", \"Amount\", \"Optional\" FROM \"Customers\" WHERE \"Country\" = ?";
    command.Parameters.Add("country", OdbcType.NVarChar, 32).Value = "Poland";
    using (var reader = command.ExecuteReader()) {
        Require(reader.FieldCount == 3 && reader.GetName(0) == "Name", "Result metadata");
        Require(reader.Read(), "First row");
        Require(reader.GetString(0) == "Alicja Żółć", "Unicode value");
        Require(reader.GetDouble(1) == 12.5, "Numeric value");
        Require(!reader.IsDBNull(2) && reader.GetString(2) == "", "Empty string");
        Require(!reader.Read(), "Parameter filters result");
    }
    var tables = connection.GetSchema("Tables");
    Require(tables.Rows.Count == 1 && (string)tables.Rows[0]["TABLE_NAME"] == "Customers", "Table discovery");
    var columns = connection.GetSchema("Columns", [null, null, "Customers", null]);
    Require(columns.Rows.Count == 6, "Column discovery");
    command.Parameters.Clear();
    command.CommandText = "SELECT \"Optional\" FROM \"Customers\" WHERE \"Name\" = 'Bob'";
    using (var reader = command.ExecuteReader()) Require(reader.Read() && reader.IsDBNull(0), "Null value");
    Console.WriteLine("System.Data.Odbc: parameterized queries, typed rows, nulls and schema discovery passed.");
} finally {
    if (Native.RegOverridePredefKey(unchecked((nint)(long)0xffffffff80000002), 0) != 0)
        throw new InvalidOperationException("Cannot restore the verification process's registry mapping.");
}

static void Require(bool condition, string contract) {
    if (!condition) throw new InvalidOperationException("ODBC contract failed: " + contract);
}

internal static partial class Native {
    [LibraryImport("advapi32.dll")]
    internal static partial int RegOverridePredefKey(nint predefined, nint replacement);
}
