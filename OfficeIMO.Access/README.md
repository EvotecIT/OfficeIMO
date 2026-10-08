# OfficeIMO.Access

OfficeIMO.Access provides typed in-memory Access models, bounded native header inspection, and explicit operation assessments. It targets Jet 4 MDB and ACE 12 ACCDB models. It does not yet decode native catalogs or write native database files.

The library runs without Microsoft Access, DAO, OLE DB or a database provider. It references OfficeIMO.Core for shared document contracts and bounded stream handling. The Windows Access/DAO tools in the verification project are independent test producers and consumers.

## Build and use the source

```sh
dotnet build OfficeIMO.Access/OfficeIMO.Access.csproj -c Release
dotnet run --project OfficeIMO.Access.Verification/OfficeIMO.Access.Verification.csproj -- OfficeIMO.Access.Tests/Fixtures
```

The verification consumer compiles the public examples and checks their actual supported behavior.

## Create and inspect a model

```csharp
using OfficeIMO.Access;

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

using (AccessDataReader reader = contacts.OpenDataReader()) {
    while (reader.Read()) {
        Console.WriteLine(reader.GetString(1));
        // False: Id was omitted. An explicit null would be specified and DBNull.
        Console.WriteLine(reader.IsSpecified(0));
    }
}

AccessOperationReport assessment = database.AssessSave("contacts.accdb");
Console.WriteLine(assessment.Status); // Unsupported: native writing is not qualified.
```

Models retain supplied values and typed references. They do not execute defaults, allocate AutoNumber values, enforce database indexes or referential integrity, execute SQL, run macros, or calculate report expressions. Query definitions are inert text. Binary values are copied when appended and when returned by a reader.

## Inspect native input

```csharp
using OfficeIMO;
using OfficeIMO.Access;

using AccessDocument source = AccessDocument.Load("source.mdb", new AccessLoadOptions {
    AccessMode = DocumentAccessMode.ReadOnly,
    MaxInputBytes = 64L * 1024 * 1024,
    MaxPages = 16_384
});
Console.WriteLine(source.Profile);
Console.WriteLine(source.Inspection!.Sha256);
Console.WriteLine(source.CatalogStatus); // NotDecoded
```

Inspection recognizes the engine and generation, checks physical page alignment and resource budgets, and hashes a bounded snapshot. It does not establish catalog validity, protection state, object absence or preservation. Native table lookup throws an explicit unsupported exception. Query, form, report, macro and VBA inventories report `NotDecoded` through the document's catalog state; their empty model collections do not mean that a native database contains no objects.

Caller-owned streams stay open. Seekable input is read from the start and its position is restored on success and failure. Native file loading releases its input handle before returning. `ValidateSourceIdentity()` detects changes to a loaded path. Readers hold a model lease that blocks mutation until they are closed. An uncommitted update rolls back on disposal; objects removed by rollback cannot be edited through retained references.

## Assess output before touching storage

`AssessSave` returns an immutable report for a document identity and revision. `RequireCurrent` rejects stale assessments; `RequireNoLoss` rejects an unavailable codec. `Save` throws `AccessOperationNotSupportedException` with that report before creating output or changing a destination stream. Allowing conversion loss never enables an unsupported writer. Read-only documents reject mutations and saving. Persistence is explicit; `SaveOnDispose` is unavailable while native writing remains unqualified.

See [support and feasibility](SUPPORT.md), the [generated operation matrix](CAPABILITIES.md), and the [single product roadmap](../Docs/ROADMAP.md#microsoft-access-document-library).
