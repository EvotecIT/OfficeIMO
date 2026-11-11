---
title: "OfficeIMO.Access"
description: "Read Access MDB and ACCDB databases, inspect queries, forms and VBA, and create Jet 4 or ACE table databases without Microsoft Access."
layout: product
meta.seo_title: "Access MDB and ACCDB reading and creation for .NET"
product_color: "#a4373a"
product_label: "Microsoft Access database library"
install: "dotnet add package OfficeIMO.Access"
nuget: "OfficeIMO.Access"
docs_url: "/docs/access/"
api_url: "/api/access/"
meta.software.name: "OfficeIMO.Access"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.nuget.org/packages/OfficeIMO.Access"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## Native Access files without Access installed

`OfficeIMO.Access` reads Jet 3/Jet 4 MDB and ACE ACCDB files directly from their native pages. It needs no Microsoft Access, DAO, OLE DB or database provider, and it never runs SQL, VBA, macros or linked-table connections.

```csharp
using OfficeIMO.Access;

using AccessDocument database = AccessDocument.Load("orders.accdb");
AccessTable orders = database.Tables["Orders"];

using AccessDataReader reader = orders.OpenDataReader();
var table = new System.Data.DataTable();
table.Load(reader);
Console.WriteLine($"{orders.Name}: {table.Rows.Count} rows");
```

## Useful for

- Migrating data out of legacy MDB and ACCDB files on Windows, Linux or macOS.
- Auditing table schemas, indexes, relationships, saved queries and linked tables.
- Inventorying forms, reports, macros and VBA modules without opening the application.
- Updating VBA modules and form or report code-behind in qualified Jet 4 and ACE files.
- Generating new Jet 4 MDB or ACE 12 ACCDB table databases with rows, keys and indexes.

## Built-in safeguards

Loading is bounded by input, page, row and value limits. Binary fields and attachments are read lazily with checksum validation. Linked-table credentials are redacted, and unsupported values are returned as exact opaque bytes with a diagnostic instead of being coerced.

## Current limits

Existing databases are read-only for rows and schema; editing is limited to VBA and qualified form/report events. Profile conversion, protected-database decryption, signature operations and authored queries, designers or macros are not supported. Unchanged databases save as exact byte copies.

Start with the [Access guide](/docs/access/) or review the [support contract](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Access/SUPPORT.md).
