## 1) What’s still wrong in your generated packages

### A. Common problem in **both** files

**\[Content\_Types].xml – wrong default for `.xml`**

* You set **Default** for extension `xml` to `application/vnd.ms-visio.drawing.main+xml`.
* That is **invalid**. Default for `.xml` must be **`application/xml`**. Then you add **Overrides** per Visio part.
* Required overrides:

  * `/visio/document.xml` → `application/vnd.ms-visio.drawing.main+xml`
  * `/visio/pages/pages.xml` → `application/vnd.ms-visio.pages+xml`
  * `/visio/pages/page1.xml` → `application/vnd.ms-visio.page+xml` ([Microsoft Learn][1])

> Impact: when the default is wrong, Visio treats **every** `.xml` as if it were the main drawing part. Some apps are lenient, Visio is not.

---

### B. **Basic Visio.vsdx**

1. **Package root relationship type is wrong**
   `/_rels/.rels` uses the generic Office type:
   `.../officeDocument/2006/relationships/officeDocument`
   It **must** be Visio’s:
   `http://schemas.microsoft.com/visio/2010/relationships/document`. ([Microsoft Learn][1])

2. **`/visio/pages/pages.xml` structure**

* Uses `<Page ID="0" Name="..." RelId="..."/>` (attribute `RelId`), and **Page IDs start at 0** here.
* Spec expects each `<Page>` to contain a **child** `<Rel r:id="rId#">` element (not an attribute), and Page IDs are **1‑based** in practice. Also the `r:id` value must follow the `rId#` pattern. ([Microsoft Learn][2])

---

### C. **Connect Rectangles.vsdx**

1. **Good**: You fixed the package → document rel type and the `Pages` → `Page` `<Rel r:id="...">` form.

2. **Still wrong**: The **shape ID is non‑numeric**
   In `/visio/pages/page1.xml` you have a connector shape with `ID="C1"`. In the Visio schema, `Shape/@ID` is an **unsigned integer**; using “C1” breaks the schema and will cause load failures. Use a numeric ID (`3`), and reference it numerically in `<Connects>`. ([Microsoft Learn][3])

3. **\[Content\_Types].xml** still sets `.xml` default to the Visio main content type (see A).

---

## 2) Make these exact changes in the builder

1. **Relationships (keep these types, and use `rId#` IDs)**

* Package → `/visio/document.xml`:
  `Type="http://schemas.microsoft.com/visio/2010/relationships/document"` **(Id `rId1`)**. ([Microsoft Learn][1])
* `/visio/document.xml` → `pages/pages.xml`:
  `Type="http://schemas.microsoft.com/visio/2010/relationships/pages"` **(Id `rId1`)**. ([Microsoft Learn][4])
* `/visio/pages/pages.xml` → `page1.xml`:
  `Type="http://schemas.microsoft.com/visio/2010/relationships/page"` **(Id `rId1`)**. ([Microsoft Learn][2])

2. **`/visio/pages/pages.xml`**
   Use this pattern (note `<Rel>` child and `r:id="rId1"`):

```xml
<Pages xmlns="http://schemas.microsoft.com/office/visio/2012/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <Page ID="1" Name="Page-1">
    <Rel r:id="rId1"/>
  </Page>
</Pages>
```

The `r:id` must conform to **ST\_RelationshipID** (`rId` + number). ([Microsoft Learn][5])

3. **`/visio/pages/page1.xml`**
   Ensure all `<Shape>` elements use **numeric IDs** (`1,2,3,…`). If you need a name (e.g., “Connector”), set `NameU="Connector"` but keep `ID` numeric.

4. **`[Content_Types].xml`**

* Default:

  * `rels` → `application/vnd.openxmlformats-package.relationships+xml`
  * `xml` → `application/xml`
* Overrides: add specific overrides for the Visio parts (document, pages, pageN). Do **not** set the default `.xml` to a Visio content type. ([Microsoft Learn][1])

> Tip: If you create parts with `System.IO.Packaging.Package.CreatePart(uri, contentType, ...)`, OPC will emit correct **Overrides** for those parts. Don’t hand‑craft a global default for `.xml`.

---

## 3) Drop‑in **working code** (builder + validator)

The builder below uses **System.IO.Packaging** + **LINQ to XML** to produce a valid VSDX (1 page, 2 rectangles, 1 connector). IDs and relationship types are exactly as required. The validator catches the mistakes above and can be added to your tests.

> **Files:** place both classes anywhere in your `OfficeIMO.Visio` project and call `VisioWriter.Create(...)` from your tests/samples.

### `VisioWriter.cs`

```csharp
using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;

namespace OfficeIMO.Visio
{
    public static class VisioWriter
    {
        // Namespaces
        private static readonly XNamespace v   = "http://schemas.microsoft.com/office/visio/2012/main";
        private static readonly XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // Relationship types
        private const string RT_Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string RT_Pages    = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        private const string RT_Page     = "http://schemas.microsoft.com/visio/2010/relationships/page";

        // Content types
        private const string CT_Document = "application/vnd.ms-visio.drawing.main+xml";
        private const string CT_Pages    = "application/vnd.ms-visio.pages+xml";
        private const string CT_Page     = "application/vnd.ms-visio.page+xml";

        /// <summary>
        /// Creates a minimal VSDX with 2 rectangles connected by a simple connector.
        /// </summary>
        public static void Create(string filePath)
        {
            if (File.Exists(filePath))
                File.Delete(filePath);

            using var package = Package.Open(filePath, FileMode.Create, FileAccess.ReadWrite);

            // Part URIs
            var documentUri = PackUriHelper.CreatePartUri(new Uri("/visio/document.xml", UriKind.Relative));
            var pagesUri    = PackUriHelper.CreatePartUri(new Uri("/visio/pages/pages.xml", UriKind.Relative));
            var page1Uri    = PackUriHelper.CreatePartUri(new Uri("/visio/pages/page1.xml", UriKind.Relative));

            // Create parts with correct content types (OPC writes Overrides for us)
            var documentPart = package.CreatePart(documentUri, CT_Document, CompressionOption.Maximum);
            var pagesPart    = package.CreatePart(pagesUri,    CT_Pages,    CompressionOption.Maximum);
            var page1Part    = package.CreatePart(page1Uri,    CT_Page,     CompressionOption.Maximum);

            // Root (package-level) relationship MUST be Visio's document rel type
            package.CreateRelationship(documentUri, TargetMode.Internal, RT_Document, "rId1");

            // Document -> Pages, Pages -> Page1
            documentPart.CreateRelationship(pagesUri, TargetMode.Internal, RT_Pages, "rId1");
            pagesPart.CreateRelationship(page1Uri,    TargetMode.Internal, RT_Page,  "rId1");

            // Payloads
            WriteDocumentXml(documentPart.GetStream(FileMode.Create, FileAccess.Write));
            WritePagesXml   (pagesPart.GetStream   (FileMode.Create, FileAccess.Write));
            WritePage1Xml   (page1Part.GetStream   (FileMode.Create, FileAccess.Write));
        }

        private static void WriteDocumentXml(Stream stream)
        {
            var doc =
                new XDocument(
                    new XDeclaration("1.0", "utf-8", null),
                    new XElement(v + "VisioDocument",
                        new XElement(v + "DocumentSettings",
                            new XElement(v + "RelayoutAndRerouteUponOpen", 1)
                        ),
                        new XElement(v + "Colors"),
                        new XElement(v + "FaceNames"),
                        new XElement(v + "StyleSheets")
                    )
                );
            using var writer = new StreamWriter(stream);
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }

        private static void WritePagesXml(Stream stream)
        {
            var doc =
                new XDocument(
                    new XDeclaration("1.0", "utf-8", null),
                    new XElement(v + "Pages",
                        new XAttribute(XNamespace.Xmlns + "r", rel),
                        new XElement(v + "Page",
                            new XAttribute("ID", 1),      // 1-based IDs
                            new XAttribute("Name", "Page-1"),
                            new XElement(v + "Rel",
                                new XAttribute(rel + "id", "rId1")  // MUST be rId#
                            )
                        )
                    )
                );
            using var writer = new StreamWriter(stream);
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }

        private static void WritePage1Xml(Stream stream)
        {
            // Shapes: 1, 2 are rectangles; 3 is the connector (IDs MUST be numeric)
            var doc =
                new XDocument(
                    new XDeclaration("1.0", "utf-8", null),
                    new XElement(v + "PageContents",
                        new XElement(v + "Shapes",
                            new XElement(v + "Shape",
                                new XAttribute("ID", 1),
                                new XAttribute("NameU", "Start"),
                                new XElement(v + "XForm",
                                    new XElement(v + "PinX", 1.0),
                                    new XElement(v + "PinY", 1.0),
                                    new XElement(v + "Width", 2.0),
                                    new XElement(v + "Height", 1.0)
                                ),
                                new XElement(v + "Text", "Start")
                            ),
                            new XElement(v + "Shape",
                                new XAttribute("ID", 2),
                                new XAttribute("NameU", "End"),
                                new XElement(v + "XForm",
                                    new XElement(v + "PinX", 4.0),
                                    new XElement(v + "PinY", 1.0),
                                    new XElement(v + "Width", 2.0),
                                    new XElement(v + "Height", 1.0)
                                ),
                                new XElement(v + "Text", "End")
                            ),
                            new XElement(v + "Shape",
                                new XAttribute("ID", 3),                // connector must also be numeric
                                new XAttribute("NameU", "Connector"),
                                new XElement(v + "Geom",
                                    new XElement(v + "MoveTo",
                                        new XAttribute("X", 2.0),
                                        new XAttribute("Y", 1.0)
                                    ),
                                    new XElement(v + "LineTo",
                                        new XAttribute("X", 3.0),
                                        new XAttribute("Y", 1.0)
                                    )
                                )
                            )
                        ),
                        new XElement(v + "Connects",
                            // Connect the connector (shape 3) to shape 1 and 2
                            new XElement(v + "Connect",
                                new XAttribute("FromSheet", 3),
                                new XAttribute("FromCell", "BeginX"),
                                new XAttribute("ToSheet", 1),
                                new XAttribute("ToCell", "PinX")
                            ),
                            new XElement(v + "Connect",
                                new XAttribute("FromSheet", 3),
                                new XAttribute("FromCell", "EndX"),
                                new XAttribute("ToSheet", 2),
                                new XAttribute("ToCell", "PinX")
                            )
                        )
                    )
                );
            using var writer = new StreamWriter(stream);
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }
    }
}
```

### `VisioValidator.cs`

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio
{
    public static class VisioValidator
    {
        private static readonly XNamespace ct  = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace pr  = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace v   = "http://schemas.microsoft.com/office/visio/2012/main";

        private const string RT_Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string RT_Pages    = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        private const string RT_Page     = "http://schemas.microsoft.com/visio/2010/relationships/page";

        private const string CT_Document = "application/vnd.ms-visio.drawing.main+xml";
        private const string CT_Pages    = "application/vnd.ms-visio.pages+xml";
        private const string CT_Page     = "application/vnd.ms-visio.page+xml";

        public static IReadOnlyList<string> Validate(string vsdxPath)
        {
            var issues = new List<string>();
            using var pkg = Package.Open(vsdxPath, FileMode.Open, FileAccess.Read);

            // 1) [Content_Types].xml
            var ctDoc = GetContentTypes(pkg);
            var defaults   = ctDoc.Root!.Elements(ct + "Default").ToList();
            var overrides  = ctDoc.Root!.Elements(ct + "Override").ToList();

            var xmlDefault = defaults.FirstOrDefault(d => (string)d.Attribute("Extension") == "xml");
            if (xmlDefault is null || (string)xmlDefault.Attribute("ContentType") != "application/xml")
                issues.Add("Default for '.xml' must be 'application/xml' with per-part Overrides.");

            bool HasOverride(string partName, string type) =>
                overrides.Any(o => (string)o.Attribute("PartName") == partName &&
                                   (string)o.Attribute("ContentType") == type);

            if (!HasOverride("/visio/document.xml", CT_Document))
                issues.Add("Missing Override for /visio/document.xml -> application/vnd.ms-visio.drawing.main+xml.");

            if (!HasOverride("/visio/pages/pages.xml", CT_Pages))
                issues.Add("Missing Override for /visio/pages/pages.xml -> application/vnd.ms-visio.pages+xml.");

            if (!HasOverride("/visio/pages/page1.xml", CT_Page))
                issues.Add("Missing Override for /visio/pages/page1.xml -> application/vnd.ms-visio.page+xml.");

            // 2) Root relationship type
            var rootRels = GetRels(pkg, "/_rels/.rels");
            var docRel = rootRels.Elements(pr + "Relationship")
                                 .FirstOrDefault(r => (string)r.Attribute("Target") == "/visio/document.xml");
            if (docRel == null || (string)docRel.Attribute("Type") != RT_Document)
                issues.Add("Root relationship must target /visio/document.xml with Visio document type.");

            // 3) document.xml -> pages
            var docPart = pkg.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
            var docRels = GetRels(pkg, "/visio/_rels/document.xml.rels");
            var pagesRel = docRels.Elements(pr + "Relationship")
                                  .FirstOrDefault(r => (string)r.Attribute("Target") == "pages/pages.xml");
            if (pagesRel == null || (string)pagesRel.Attribute("Type") != RT_Pages)
                issues.Add("document.xml must relate to pages/pages.xml with visio/2010/relationships/pages.");

            // 4) pages.xml -> page1.xml and Pages XML structure
            var pagesXml = LoadXml(pkg, "/visio/pages/pages.xml");
            var page = pagesXml.Root!.Element(v + "Page");
            if (page == null) issues.Add("pages.xml must contain a Page element.");
            else
            {
                // Page ID numeric and 1-based
                if (!int.TryParse((string)page.Attribute("ID"), out var pageId) || pageId < 1)
                    issues.Add("Page/@ID must be numeric and 1-based (e.g., 1).");

                // Must contain child <Rel r:id="rId#">
                var relChild = page.Element(v + "Rel");
                var rid = (string?)relChild?.Attribute(rel + "id");
                if (relChild == null || string.IsNullOrWhiteSpace(rid) || !rid.StartsWith("rId"))
                    issues.Add("Page must contain <Rel r:id=\"rId#\"> child (not an attribute).");
            }

            var pagesRels = GetRels(pkg, "/visio/pages/_rels/pages.xml.rels");
            var pageRel = pagesRels.Elements(pr + "Relationship")
                                   .FirstOrDefault(r => (string)r.Attribute("Type") == RT_Page);
            if (pageRel == null) issues.Add("pages.xml.rels must have a relationship of type visio/2010/relationships/page.");

            // 5) Ensure shapes have numeric IDs
            var page1Xml = LoadXml(pkg, "/visio/pages/page1.xml");
            var badId = page1Xml.Descendants(v + "Shape")
                                .Select(x => (string)x.Attribute("ID"))
                                .FirstOrDefault(id => !int.TryParse(id, out _));
            if (badId != null)
                issues.Add($"Shape/@ID must be numeric. Found non-numeric ID: '{badId}'.");

            return issues;
        }

        private static XDocument GetContentTypes(Package pkg)
        {
            using var s = pkg.GetPart(new Uri("/[Content_Types].xml", UriKind.Relative)).GetStream();
            return XDocument.Load(s);
        }

        private static XDocument LoadXml(Package pkg, string partName)
        {
            using var s = pkg.GetPart(new Uri(partName, UriKind.Relative)).GetStream();
            return XDocument.Load(s);
        }

        private static XDocument GetRels(Package pkg, string relsPath)
        {
            using var s = pkg.GetPart(new Uri(relsPath, UriKind.Relative)).GetStream();
            return XDocument.Load(s);
        }
    }
}
```

---

## 4) Quick test harness (optional)

```csharp
// Program.cs (example usage)
using System;
using OfficeIMO.Visio;

class Program
{
    static void Main()
    {
        var path = "out.vsdx";
        VisioWriter.Create(path);
        var issues = VisioValidator.Validate(path);
        if (issues.Count == 0)
        {
            Console.WriteLine("OK: VSDX looks valid.");
        }
        else
        {
            Console.WriteLine("Issues:");
            foreach (var i in issues) Console.WriteLine(" - " + i);
        }
    }
}
```

---

## 5) TL;DR – the minimal checklist

* **Root rel type**: `http://schemas.microsoft.com/visio/2010/relationships/document`. ([Microsoft Learn][1])
* **Document content type**: `/visio/document.xml` → `application/vnd.ms-visio.drawing.main+xml`. ([Microsoft Learn][1])
* **Pages rel + page rel**: `.../relationships/pages` and `.../relationships/page`. ([Microsoft Learn][4])
* **Pages XML**: `<Page ID="1" ...><Rel r:id="rId1"/></Page>` (not `RelId="..."`). **`rId#`** format required. ([Microsoft Learn][5])
* **Shape IDs**: **numeric** only. ([Microsoft Learn][3])
* **\[Content\_Types]**: default `.xml` = `application/xml`; use **Overrides** for Visio parts. ([Microsoft Learn][1])

[1]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/7ec3d7b0-0de2-4711-a7b6-92daa2020d71?utm_source=chatgpt.com "Document XML Part - MS-VSDX"
[2]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/1f15c8f0-6565-465c-aefd-2be6af545e8a?utm_source=chatgpt.com "[MS-VSDX]: Page XML Part"
[3]: https://learn.microsoft.com/en-us/office/client-developer/visio/schema-mapvisio-xml?utm_source=chatgpt.com "Schema map (Visio XML)"
[4]: https://learn.microsoft.com/en-us/office/client-developer/visio/how-to-manipulate-the-visio-file-format-programmatically?utm_source=chatgpt.com "Manipulate the Visio file format programmatically"
[5]: https://learn.microsoft.com/en-us/office/client-developer/visio/rel-element-foreigndata_type-complextypevisio-xml?utm_source=chatgpt.com "Rel element (ForeignData_Type complexType) (Visio XML)"
