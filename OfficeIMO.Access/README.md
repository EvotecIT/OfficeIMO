# OfficeIMO.Access

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Access)](https://www.nuget.org/packages/OfficeIMO.Access)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Access?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Access)

Read Access databases and create new ones in .NET without Microsoft Access, DAO, OLE DB or a database provider. `OfficeIMO.Access` reads Jet 3 and Jet 4 `.mdb` and ACE `.accdb` files, creates fresh Jet 4 and ACE 12 table databases, and inspects and edits VBA projects, through one typed `AccessDocument`. Reading never runs SQL, VBA, macros or linked-table connections.

## Install

```sh
dotnet add package OfficeIMO.Access
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/Access/Program.cs" region="QuickStart" -->
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
database.AssessSave("Contacts.accdb").ThrowIfLoss();
database.Save("Contacts.accdb");
```
<!-- /powerforge:code -->

`AssessSave` reports anything the target format can't carry before the file is written.

## What it does

- **Read tables** with schemas, properties, indexes, relationships and forward-only rows, including binary and structured fields.
- **Inspect** saved queries, linked tables, forms, reports, macros and modules without running them.
- **Edit VBA** projects in qualified Jet 4 and ACE databases.
- **Create** new `.mdb` and `.accdb` table databases, reproducibly, without a seed file.
- **Preserve** an unchanged database when you save it again.
- **Strict loss policy** by default: operations report and refuse what they can't carry.

The [Access guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.access-guide.md) shows each of these with examples.

## Limits

Execution of SQL and VBA is out of scope; [DbaClientX](https://github.com/EvotecIT/DbaClientX) can execute queries where that is wanted. Reading, preservation, editing, creation and conversion are qualified separately in the [operation matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Access/SUPPORT.md), and the [Access design](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.access-design.md) explains the model.

## Learn more

- [OfficeIMO.Access on officeimo.com](https://officeimo.com/products/access/) and the [Access docs](https://officeimo.com/docs/access/).
- [Access guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.access-guide.md) and the [operation matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Access/SUPPORT.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` only.
- License: MIT.

