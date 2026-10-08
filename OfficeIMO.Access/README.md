# OfficeIMO.Access

OfficeIMO.Access reads native Jet 4 MDB and ACE ACCDB files through one typed document model. It exposes catalogs, selected table schemas and properties, index definitions, relationships, saved-query records and forward-only rows. Attachments, multivalued fields and long binary values have bounded lazy access.

The library runs on .NET without Microsoft Access, DAO, OLE DB or a database provider. It references OfficeIMO.Core for shared document contracts and binary primitives. File inspection never executes SQL, VBA, macros, calculated expressions or linked-table connections.

## Build and use the source

```sh
dotnet build OfficeIMO.Access/OfficeIMO.Access.csproj -c Release
dotnet run --project OfficeIMO.Access.Verification/OfficeIMO.Access.Verification.csproj -- OfficeIMO.Access.Tests/Fixtures
```

The verification executable compiles the public examples and checks them against synthetic files independently produced and consumed by Microsoft Access/DAO. The library targets `netstandard2.0`, `net8.0`, `net10.0` and Windows `net472`.

## Read a native table

```csharp
using OfficeIMO;
using OfficeIMO.Access;

using AccessDocument source = AccessDocument.Load("source.mdb", new AccessLoadOptions {
    AccessMode = DocumentAccessMode.ReadOnly,
    TableNames = new[] { "Contacts" },
    MaxInputBytes = 64L * 1024 * 1024,
    MaxRows = 1_000_000
});

Console.WriteLine(source.Profile);
Console.WriteLine(source.CatalogStatus); // Decoded for qualified, unprotected input
AccessTable contacts = source.Tables["Contacts"];
foreach (AccessColumn column in contacts.Columns) {
    Console.WriteLine($"{column.Name}: {column.DataType}");
}

using AccessDataReader reader = contacts.OpenDataReader();
int name = reader.GetOrdinal("DisplayName");
while (reader.Read()) {
    Console.WriteLine(reader.IsDBNull(name) ? "(null)" : reader.GetString(name));
}
```

The same API reads both file families. `AccessDataReader` implements `DbDataReader` and works with `DataTable.Load`. Field order and types come from the native schema. `RowCount` is the declared native count; a complete scan validates it. Linked tables expose `LinkedTable` metadata and reject row access because the library never resolves external targets.

`Tables` contains user tables. `SystemTables` exposes native system/backing tables separately, and `Catalog` retains typed object identifiers, flags, owner metadata and exact catalog-row representations. Each collection has its own `CatalogStatus`: decoded table metadata does not imply decoded form, report, macro or VBA payloads. These application collections remain `NotDecoded`.

## Read binary and structured fields

```csharp
using AccessDocument source = AccessDocument.Load("values.accdb");
using AccessDataReader rows = source.Tables["Scalars"].OpenDataReader();
if (rows.Read()) {
    using Stream payload = rows.GetStream(rows.GetOrdinal("Payload"));
    // Copies incrementally from the native long-value chain.
    using Stream destination = File.Create("payload.bin");
    await payload.CopyToAsync(destination);
}

using AccessDataReader structured = source.Tables["Structured"].OpenDataReader();
if (structured.Read()) {
    var tags = (AccessComplexValue)structured["Tags"];
    foreach (object? value in tags.EnumerateValues()) Console.WriteLine(value);

    var files = (AccessComplexValue)structured["Files"];
    foreach (AccessAttachment attachment in files.EnumerateAttachments()) {
        Console.WriteLine($"{attachment.FileName}: {attachment.FileType}");
        byte[] content = attachment.GetBytes(); // Bounded, checksum-validated decode
        byte[] native = attachment.GetEncodedBytes(); // Exact encoded payload copy
    }
}
```

Ordinary fields decode when requested. Binary `GetStream` reads forward through native storage without allocating the whole payload; `GetValue` returns a defensive byte array. Attachment enumeration reads metadata first; requesting content allocates a bounded decoded payload. Structured fields also offer `OpenDataReader` for their backing rows. Complex-value scans use the same per-reader row limit.

Calculated fields, unknown representations and Decimal definitions outside the qualified .NET precision range return `AccessOpaqueValue` with a reason and column diagnostics. `GetBytes()` returns their exact representation as a defensive copy. Rich-text markup and hyperlink strings retain their persisted content; the reader does not render or follow them.

## Inspect saved queries and links

```csharp
foreach (AccessQueryDefinition query in source.Queries) {
    if (query.HasSql) Console.WriteLine(query.Sql);
    else Console.WriteLine($"{query.Name}: {query.NativeRecords.Count} native records");
}
foreach (AccessTable table in source.Tables.Where(x => x.IsLinked)) {
    Console.WriteLine(table.LinkedTable!.ForeignTableName);
    Console.WriteLine(table.LinkedTable.Connection); // Credential fields are redacted
}
```

Saved queries are inert. SQL reconstruction covers the qualified simple single-table SELECT and two-part UNION shapes. Sized parameters, external-source qualifiers and other unqualified shapes have `HasSql` false and `Sql` then throws. `NativeRecords` retains all query attributes and exact record bytes, including unsupported definitions. Original expressions and typed parameters remain available without execution.

Ordinary linked-table views and the `MSysObjects.Connect` reader redact credential fields. Explicit raw catalog-row access can contain sensitive native data; callers control whether to inspect or export those bytes.

## Inspection, limits and lifetime

`AccessDocument.Inspect(path)` performs header-only inspection. `Load` decodes qualified catalogs by default; `DecodeCatalog = false` keeps header-only loading. Protected and Jet 3 inputs retain `NotDecoded` catalogs with explicit diagnostics. Header inspection checks signatures, generation, page alignment and byte/page budgets and hashes a bounded snapshot. It does not establish protection state, object absence or whole-database validity.

Loading snapshots the bounded file into memory and closes its file handle. Table selection limits user-schema decoding, while required system metadata is still decoded. User rows and large fields are not materialized into the document model. Limits cover input bytes, pages, catalog objects, total metadata bytes, value bytes, rows traversed and chain depth. Cancellation reaches loading, page traversal and payload reads; in-memory reader async methods complete synchronously.

Caller-owned streams stay open. Seekable input is read from the start and its position is restored on success and failure. `ValidateSourceIdentity()` detects changes to a loaded path. Native documents are immutable. Keep the document open while using readers, streams, structured values or attachments; disposing it invalidates those views.

## Create an in-memory model and assess output

```csharp
using AccessDocument database = AccessDocument.Create(new AccessCreateOptions {
    Format = AccessFileFormat.Accdb
});
AccessTable contacts = database.Tables.Add("Contacts");
contacts.Columns.Add("Id", AccessDataType.AutoNumber);
contacts.Columns.Add("Name", AccessDataType.ShortText, maxLength: 120);
contacts.Indexes.AddPrimaryKey("PK_Contacts", "Id");

using (AccessUpdateScope edit = database.BeginUpdate()) {
    contacts.AppendRow(new AccessRowValues { ["Name"] = "Ada" });
    edit.Commit();
}
AccessOperationReport assessment = database.AssessSave("contacts.accdb");
Console.WriteLine(assessment.Status); // Unsupported: no production native writer
```

Models retain supplied values and typed references. They do not allocate AutoNumber values, apply defaults, enforce indexes or referential integrity, calculate expressions or execute queries. Modeled rows distinguish an omitted value from an explicit null. Reader leases block mutation; an uncommitted update rolls back on disposal.

`AssessSave` identifies known profile losses, including complex fields, rich text, calculated fields, Large Number and Date/Time Extended. It returns an immutable report tied to the document identity and revision. Native creation, editing, preservation, conversion and saving remain unsupported. `Save` fails with its report before creating or changing output. Allowing conversion loss does not enable an unavailable writer; `SaveOnDispose` is unavailable.

See [support and independent evidence](SUPPORT.md), the [generated operation matrix](CAPABILITIES.md), and the [product roadmap](../Docs/ROADMAP.md#microsoft-access-document-library).
