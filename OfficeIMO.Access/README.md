# OfficeIMO.Access

OfficeIMO.Access reads native Jet 4 MDB and ACE ACCDB files, inspects application objects, edits qualified native VBA projects, preserves unchanged files, and creates fresh Jet 4/ACE 12 table databases from one typed document model. It exposes catalogs, selected table schemas and properties, indexes, relationships, saved-query records and forward-only rows. Attachments, multivalued fields and long binary values have bounded lazy access.

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

`Tables` contains user tables. `SystemTables` exposes native system/backing tables separately, and `Catalog` retains native identifiers, flags, owner metadata and exact catalog-row representations. Each collection has its own `CatalogStatus`. An application inventory can be decoded while individual definitions remain preserve-only.

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

Caller-owned streams stay open. Seekable input is read from the start and its position is restored on success and failure. `ValidateSourceIdentity()` detects changes to a loaded path. Native row and schema editing remain unavailable; VBA uses the explicit workflow below. Keep the document open while using readers, streams, structured values or attachments; disposing it invalidates those views.

## Inspect application objects and VBA

```csharp
using AccessDocument source = AccessDocument.Load("application.accdb");
foreach (AccessApplicationObject form in source.Forms) {
    Console.WriteLine($"{form.Name}: {form.Definition?.RecordSource}");
    foreach (AccessStorageStream stream in form.Streams)
        Console.WriteLine($"{stream.Path}: {stream.Payload.GetBytes().Length} bytes");
}
foreach (AccessVbaModuleInfo module in source.VbaProject.Modules)
    Console.WriteLine(module.Source ?? module.SourceLimitation);
foreach (AccessVbaReferenceInfo reference in source.VbaProject.References)
    Console.WriteLine($"{reference.Name}: {reference.LibraryId}");
```

ACE version-21 designers expose qualified sections, controls, sources and Open, Click and AfterUpdate bindings. Jet version-19 designers and unsupported properties retain exact bytes. Standalone and embedded `StopMacro` definitions, table data-macro XML and resource metadata have separate APIs. Stored VBA references are inventoried without resolving libraries; implicit built-in references are outside this inventory. Missing compiled-source text stays unavailable.

`ApplicationStreams` retains storage paths and exact payloads, `Dependencies` contains observed references with explicit unresolved entries, and `ChangeJournal` records model mutations and rolls them back with the edit scope. Dependency inspection is partial; it does not parse arbitrary SQL or VBA. Set `DecodeApplicationObjects = false` when only the table catalog is needed.

## Edit native VBA

```csharp
using OfficeIMO;
using OfficeIMO.Access;

using AccessDocument database = AccessDocument.Load("application.accdb");
OfficeVbaProject project = database.GetVbaProject();
project.SetModuleSource("BusinessLogic",
    "Option Explicit\r\nPublic Function IsReady() As Boolean\r\n IsReady = True\r\nEnd Function\r\n");
project.AddModule("Utilities", "Public Const RetryCount As Long = 3\r\n");
database.SetVbaProject(project);
database.Save("edited.accdb");
```

`GetVbaProject` returns a detached shared Core model. Its edits become database changes only after `SetVbaProject`; persistence requires `Save`. Standard and standalone class modules support source replacement, addition, removal and rename. An existing native application without VBA accepts a project from `OfficeVbaProject.Create`. Missing or opaque source and unsupported storage layouts fail before database mutation.

The executable public workflow is `dotnet run --project OfficeIMO.Access.Verification -- --edit-vba OfficeIMO.Access.Tests/Fixtures <new-output-folder>`. On a Windows machine with Access installed, `OfficeIMO.Access.Verification/Test-NativeVbaEditing.ps1 -OutputRoot <output-folder>` independently reopens and compiles those synthetic outputs with macros disabled.

Existing form/report code-behind is available through its native module name, such as `Form_Orders` or `Report_Invoices`, with `OfficeVbaModuleKind.Document`. Replace its source with `SetModuleSource`; its native `VB_Base` and host binding are preserved. `SetCodeBehind` creates or replaces code-behind on an existing native form or report, including its first class module. A new class requires an ASCII VBA identifier of at most 31 characters including `Form_` or `Report_`; other object names reject creation. Host module rename/removal remains unqualified. Its `OfficeVbaWriteOptions` limits apply to both the existing project load and the staged output.

```csharp
AccessApplicationObject form = database.Forms["Orders"];
form.SetCodeBehind("Option Explicit\r\nPrivate Sub Form_Open(Cancel As Integer)\r\nEnd Sub\r\n");
form.SetEventBinding(AccessEventKind.Open, "[Event Procedure]");
form.SetEventBinding(AccessEventKind.AfterUpdate, "=Len(\"inert\")", "CustomerChoice");
database.Save("orders-edited.accdb");
```

Event authoring is qualified for ACE expanded version-21 designers: form/report Open, label/text-box/combo-box Click, and text-box/combo-box AfterUpdate. Bindings remain inert; setting an event does not create or run a handler. `[Event Procedure]` requires existing code-behind. Null or empty clears a binding. Replacing Click also removes its associated embedded macro; unrelated native records remain exact. Jet version-19 designers, nonempty designer deltas, other control/event layouts and new embedded actions reject event edits.

The executable host workflow is `dotnet run --project OfficeIMO.Access.Verification -- --edit-vba-hosts OfficeIMO.Access.Tests/Fixtures <new-output-folder>`. `Test-NativeVbaHosts.ps1 -OutputRoot <output-folder>` and `Test-NativeVbaHostEvents.ps1 -OutputRoot <output-folder>` independently verify native class source, event state and VBE compilation with macros disabled.

Native editing is qualified for unprotected Jet 4 and ACE 12/14 application storage. Module catalog IDs, storage slots, permission records and unrelated application payloads remain preserved. Unknown permission/index/signature carriers reject editing; signature validation and signed-project authoring are unavailable. The default `MaximumRecoveryBytes` is 64 MiB. Old unreachable pages remain allocated; this operation does not compact the database.

Close readers before applying a project. Update scopes roll back staged VBA, inventory and catalog changes together. Saving over the retained source with an explicit replacement policy uses a guarded atomic commit; an external change during staging is preserved and causes the save to fail.

Detached module identities survive intervening applications of other detached projects. Applying a module whose native identity was deleted fails before mutation, even if a new module now uses its old name; reload the current project before continuing that edit.

## Create a native database

```csharp
using AccessDocument database = AccessDocument.Create(new AccessCreateOptions {
    Format = AccessFileFormat.Accdb,
    DatabaseTitle = "Contacts"
});
AccessTable contacts = database.Tables.Add("Contacts");
contacts.Columns.AddAutoNumber("Id", seed: 101);
contacts.Columns.Add("Name", AccessDataType.ShortText, maxLength: 120);
contacts.Indexes.AddPrimaryKey("PK_Contacts", "Id");

using (AccessUpdateScope edit = database.BeginUpdate()) {
    contacts.AppendRow(new AccessRowValues { ["Name"] = "Ada" });
    edit.Commit();
}
AccessOperationReport assessment = database.AssessSave("contacts.accdb");
assessment.RequireNoLoss();
database.Save("contacts.accdb");
```

Use `Format = AccessFileFormat.Mdb` and an `.mdb` path to create Jet 4 output. Creation writes native headers, catalogs, allocation maps, table definitions, rows, long values and indexes directly; it reads no blank database or template and invokes no external engine.

Models retain supplied values and typed references. Modeled rows distinguish omission from explicit null. Saving allocates omitted sequential AutoNumber values in the output without modifying model rows, checks primary/unique keys and enforced single-field relationships, and persists omitted Yes/No as false. Explicit null Yes/No or AutoNumber values are rejected. Reader leases block mutation; an uncommitted update rolls back on disposal.

`AddDecimal(name, precision, scale)` declares exact Decimal storage without rounding. `AddUnique(name, columns)` and `Add(name, columns)` create unique and ordinary indexes. Creation supports Byte, Int16, Int32/AutoNumber and a bounded General legacy ASCII text subset as index keys. Unicode text values are supported; indexed Unicode outside that subset is rejected. Tables support up to 255 columns, 32 indexes including relationship indexes, and a 4060-byte row. The default `MaxOutputBytes` is 64 MiB. See [the creation limits](SUPPORT.md#native-creation-contract) before choosing schemas.

## Preserve an unchanged database

```csharp
using AccessDocument source = AccessDocument.Load("application.mdb");
source.AssessSave("archive.mdb").RequireNoLoss();
source.Save("archive.mdb");
```

Same-profile preservation copies the immutable loaded snapshot byte for byte, including opaque application, compiled, signed and protected content. It does not validate signatures or decode protected content. `ValidateSourceIdentity` runs before saving a loaded path. Path output uses an atomic staged commit and defaults to `FailIfExists`; pass `FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace` to replace a destination. Caller streams remain open; seekable output is rewound and truncated. Unsupported output and pre-cancellation leave the destination untouched; arbitrary stream I/O failures cannot provide file-style rollback.

`AssessSave` returns an immutable report tied to the document identity and revision and checks native output limits before destination I/O. It names known profile losses for complex fields, rich text, calculated fields, Large Number and Date/Time Extended. General native row/schema editing, profile conversion, authored queries/designers/macros and protection/signature operations remain unsupported. Allowing conversion loss does not enable an unavailable codec. Persistence is explicit; `SaveOnDispose` is unavailable.

See [support and independent evidence](SUPPORT.md), the [generated operation matrix](CAPABILITIES.md), and [remaining VBA host integration](../Docs/ROADMAP.md#vba-host-integration).
