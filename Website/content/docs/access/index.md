---
title: "Access Databases"
description: "Read MDB and ACCDB tables, inspect queries, forms and VBA, edit qualified VBA projects, and create Jet 4 or ACE databases in .NET."
order: 54
meta.seo_title: "Read and create Access MDB and ACCDB files in .NET"
verified: 2026-11-14
---

`OfficeIMO.Access` reads native Microsoft Access databases through one typed `AccessDocument` model. It decodes catalogs, table schemas, indexes, relationships, saved-query records and forward-only rows for Jet 3/Jet 4 MDB and ACE ACCDB files, inspects application objects and VBA, edits qualified VBA projects, and creates new Jet 4 or ACE 12 table databases. Microsoft Access, DAO, OLE DB and database providers are not required.

File operations are inert: the library never executes SQL, VBA, macros, calculated expressions or linked-table connections.

## Install

```shell
dotnet add package OfficeIMO.Access
```

The package targets `netstandard2.0`, `net8.0`, `net10.0` and `net472` on Windows.

## Read a native table

```csharp
using OfficeIMO;
using OfficeIMO.Access;

using AccessDocument source = AccessDocument.Load("source.accdb", new AccessLoadOptions {
    AccessMode = DocumentAccessMode.ReadOnly,
    TableNames = new[] { "Contacts" },
    MaxInputBytes = 64L * 1024 * 1024,
    MaxRows = 1_000_000
});

Console.WriteLine($"{source.Profile}: {source.CatalogStatus}");
AccessTable contacts = source.Tables["Contacts"];
foreach (AccessColumn column in contacts.Columns)
    Console.WriteLine($"{column.Name}: {column.DataType}");

using AccessDataReader reader = contacts.OpenDataReader();
int name = reader.GetOrdinal("DisplayName");
while (reader.Read())
    Console.WriteLine(reader.IsDBNull(name) ? "(null)" : reader.GetString(name));
```

The same calls read MDB and ACCDB files. `AccessDataReader` is a `DbDataReader`, so it also works with `DataTable.Load`. Field order and types come from the native schema. Binary and long values stream from native storage through `GetStream`; attachments and multivalued fields are exposed as `AccessComplexValue` with bounded, checksum-validated decoding.

`TableNames` limits user-schema decoding; required system metadata is still read. Load options also bound input bytes, pages, catalog objects, metadata bytes, value bytes, rows traversed and chain depth. Keep the document open while readers, streams or attachments are in use.

## Inspect queries, links and application objects

```csharp
using OfficeIMO.Access;

using AccessDocument database = AccessDocument.Load("application.accdb");

foreach (AccessQueryDefinition query in database.Queries)
    Console.WriteLine(query.HasSql ? query.Sql : $"{query.Name}: {query.NativeRecords.Count} native records");

foreach (AccessTable table in database.Tables.Where(t => t.IsLinked))
    Console.WriteLine($"{table.Name} -> {table.LinkedTable!.ForeignTableName}");

foreach (AccessApplicationObject form in database.Forms)
    Console.WriteLine($"{form.Name}: {form.Definition?.RecordSource}");

foreach (AccessVbaModuleInfo module in database.VbaProject.Modules)
    Console.WriteLine(module.Source ?? module.SourceLimitation);
```

Saved queries are inert records. SQL text is reconstructed for qualified simple single-table `SELECT` and two-part `UNION` shapes; other definitions report `HasSql = false` and keep their exact native records. Linked-table connection strings are shown with credential fields redacted, and linked targets are never opened.

## Edit VBA

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

`GetVbaProject` returns a detached model; changes reach the database only through `SetVbaProject` and `Save`. Standard and class modules support replacement, addition, removal and rename, and existing form or report code-behind can be replaced. VBA editing is qualified for unprotected Jet 4 and ACE 12/14 application storage. Signed projects cannot be authored, and the save does not compact the database.

## Create a database

```csharp
using OfficeIMO.Access;

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

database.AssessSave("contacts.accdb").ThrowIfLoss();
database.Save("contacts.accdb");
```

Use `AccessFileFormat.Mdb` with an `.mdb` path for Jet 4 output. Creation writes headers, catalogs, allocation maps, table definitions, rows, long values and indexes directly; it does not copy a blank template or call an external engine. Saving checks primary and unique keys and enforced single-field relationships. Set `CreatedAt` for repeatable output bytes.

## Preserve an unchanged file

`Save` on a loaded, unmodified database copies the loaded snapshot byte for byte, including application, compiled, signed and protected content. Path saves use an atomic staged commit and refuse to overwrite an existing file unless `FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace` is set. `AssessSave` reports known profile losses before any output is written.

## Limits

- General native row and schema editing of existing databases is not supported; existing files accept VBA and qualified form/report event edits only.
- Profile conversion (for example MDB to ACCDB), authored queries, form/report designers and macros, and protection or signature operations are not supported.
- Protected and encrypted databases are not decrypted; their catalogs remain undecoded with diagnostics.
- Jet 3 (Access 97) files are read-only for catalog, schema and rows; their property maps and application definitions are preserved but not decoded, and VBA inspection is unavailable.
- Calculated fields, unsupported encodings and out-of-range Decimal definitions return `AccessOpaqueValue` with the exact stored bytes and a reason.
- Header-only `AccessDocument.Inspect` does not establish protection state or whole-database validity.

## Related

- [Access support contract](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Access/SUPPORT.md) with qualified profiles and independent evidence
- [Generated operation matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Access/CAPABILITIES.md)
- [API reference](/api/access/)
- [OfficeIMO.Access product page](/products/access/)
