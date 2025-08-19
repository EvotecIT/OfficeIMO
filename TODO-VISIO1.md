## What’s still wrong (precise findings)

### 1) `[Content_Types].xml` is still incorrect in **all three** files

* You set the **Default** for extension `xml` to
  `application/vnd.ms-visio.drawing.main+xml`.
* There is **no Override** for `/visio/document.xml`.

Observed (all three files):

```
Defaults:
  - xml  => application/vnd.ms-visio.drawing.main+xml
  - rels => application/vnd.openxmlformats-package.relationships+xml
Overrides:
  - /visio/pages/pages.xml  => application/vnd.ms-visio.pages+xml
  - /visio/pages/page1.xml  => application/vnd.ms-visio.page+xml
```

What Visio expects (and what the spec actually describes by part):

* **Document part** (`/visio/document.xml`) content type:
  `application/vnd.ms-visio.drawing.main+xml`. ([Microsoft Learn][1])
* **Pages part** (`/visio/pages/pages.xml`) content type:
  `application/vnd.ms-visio.pages+xml`. ([Microsoft Learn][2])
* **Page part** (`/visio/pages/pageN.xml`) content type:
  `application/vnd.ms-visio.page+xml`. ([Microsoft Learn][2])

> The robust way to represent this in OPC is:
> **Default** for `.xml` → `application/xml`, and **Overrides** per Visio part above. (That’s exactly how OOXML/OPC packages are modeled: **Default** maps an extension, **Override** maps a **specific part name**. ) ([c-rex.net][3])

This “global default to Visio main” is the blocker that still makes Visio reject the files—even when the relationships are now correct.

---

### 2) Relationship graph (now mostly correct)

* **Package → Document**: `/_rels/.rels`
  `Type="http://schemas.microsoft.com/visio/2010/relationships/document"` ✓ (all 3)
* **Document → Pages**: `/visio/_rels/document.xml.rels`
  `Type="http://schemas.microsoft.com/visio/2010/relationships/pages"` ✓ (all 3)
* **Pages → Page1**: `/visio/pages/_rels/pages.xml.rels`
  `Type="http://schemas.microsoft.com/visio/2010/relationships/page"` ✓ (all 3)

These types match Microsoft’s schema. ([Microsoft Learn][1])

Minor hygiene:

* In **Validated.vsdx** the targets are **absolute** (`/visio/pages/pages.xml`). Prefer **relative** targets (`pages/pages.xml`, `page1.xml`) from the source part; this is how Office emits them and it avoids edge‑case URI resolution. (Not a hard fail, but make it consistent with Word/Excel/Visio packages.) ([Wikipedia][4])

---

### 3) Page + shapes XML

* **Basic Visio.vsdx**: one rectangle with numeric ID (`1`) → OK.
* **Connect Rectangles.vsdx**: has a connector with **ID="C1"** → **invalid**; shape IDs must be **integers**. Use `3` and reference it numerically in `<Connects>`. (This alone will not corrupt loading if content types are fixed, but it violates the schema.) ([Microsoft Learn][2])
* **Validated.vsdx**: `ID="1"`, `ID="2"`, `ID="3"` → OK.

The root element names and namespaces are fine:

* `/visio/document.xml` → `<VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main">…` ✓ (that’s the correct root for `document.xml`). ([Microsoft Learn][5])
* `/visio/pages/page1.xml` → `<PageContents …>` ✓ required for a Page part. ([Microsoft Learn][2])

---

## Fix it at the source (OfficeIMO.Visio)

There are two safe patterns. Pick exactly one and stick to it.

### **Pattern A — Let OPC write `[Content_Types].xml` automatically (recommended)**

Do **not** write `[Content_Types].xml` yourself.
Create parts with the correct content types and OPC will emit the proper **Overrides**:

```csharp
// Content types
const string CT_Document = "application/vnd.ms-visio.drawing.main+xml";
const string CT_Pages    = "application/vnd.ms-visio.pages+xml";
const string CT_Page     = "application/vnd.ms-visio.page+xml";

// URIs
var documentUri = PackUriHelper.CreatePartUri(new Uri("/visio/document.xml", UriKind.Relative));
var pagesUri    = PackUriHelper.CreatePartUri(new Uri("/visio/pages/pages.xml", UriKind.Relative));
var page1Uri    = PackUriHelper.CreatePartUri(new Uri("/visio/pages/page1.xml", UriKind.Relative));

using var package = Package.Open(filePath, FileMode.Create, FileAccess.ReadWrite);

// Create parts with correct types (OPC will add proper <Override/> entries)
var documentPart = package.CreatePart(documentUri, CT_Document, CompressionOption.Maximum);
var pagesPart    = package.CreatePart(pagesUri,    CT_Pages,    CompressionOption.Maximum);
var page1Part    = package.CreatePart(page1Uri,    CT_Page,     CompressionOption.Maximum);

// Relationships (exact URIs and rIds)
const string RT_Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
const string RT_Pages    = "http://schemas.microsoft.com/visio/2010/relationships/pages";
const string RT_Page     = "http://schemas.microsoft.com/visio/2010/relationships/page";

package.CreateRelationship(documentUri, TargetMode.Internal, RT_Document, "rId1");
documentPart.CreateRelationship(pagesUri, TargetMode.Internal, RT_Pages, "rId1");
pagesPart.CreateRelationship(page1Uri, TargetMode.Internal, RT_Page,  "rId1");

// Write payloads (VisioDocument, Pages with <Rel r:id="rId1"/>, PageContents…)
```

This yields:

```
Default:
  .rels -> application/vnd.openxmlformats-package.relationships+xml
  .xml  -> application/xml        // by default, provided by OPC
Overrides:
  /visio/document.xml -> application/vnd.ms-visio.drawing.main+xml
  /visio/pages/pages.xml -> application/vnd.ms-visio.pages+xml
  /visio/pages/page1.xml -> application/vnd.ms-visio.page+xml
```

### **Pattern B — If you insist on writing `[Content_Types].xml` yourself**

Write **exactly** this file:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>

  <Override PartName="/visio/document.xml"     ContentType="application/vnd.ms-visio.drawing.main+xml"/>
  <Override PartName="/visio/pages/pages.xml"  ContentType="application/vnd.ms-visio.pages+xml"/>
  <Override PartName="/visio/pages/page1.xml"  ContentType="application/vnd.ms-visio.page+xml"/>
</Types>
```

(Again: **Default** for `.xml` is `application/xml`; the three Visio parts are added via **Override**.) ([c-rex.net][3])

---

## Minimal, working `Pages` & `Page` XML (for reference)

`/visio/pages/pages.xml`:

```xml
<Pages xmlns="http://schemas.microsoft.com/office/visio/2012/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <Page ID="1" Name="Page-1">
    <Rel r:id="rId1"/>
  </Page>
</Pages>
```

`/visio/pages/page1.xml` (IDs must be **numeric**):

```xml
<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main">
  <Shapes>
    <Shape ID="1" NameU="Start">
      <XForm><PinX>1</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm>
      <Text>Start</Text>
    </Shape>
    <Shape ID="2" NameU="End">
      <XForm><PinX>4</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm>
      <Text>End</Text>
    </Shape>
    <Shape ID="3" NameU="Connector">
      <Geom>
        <MoveTo X="2" Y="1"/><LineTo X="3" Y="1"/>
      </Geom>
    </Shape>
  </Shapes>
  <Connects>
    <Connect FromSheet="3" FromCell="BeginX" ToSheet="1" ToCell="PinX"/>
    <Connect FromSheet="3" FromCell="EndX"   ToSheet="2" ToCell="PinX"/>
  </Connects>
</PageContents>
```

* Page XML must be the target of a **Pages** relationship of type `.../relationships/page`, and its root is **`PageContents`**. ([Microsoft Learn][2])
* `VisioDocument` root for `document.xml` is in namespace `http://schemas.microsoft.com/office/visio/2012/main`. ([Microsoft Learn][5])

---

## Drop‑in validator (add to your tests)

This catches exactly what broke your packages—**including** an explicit check for the bad default in `[Content_Types].xml` and the missing document override.

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

public static class VsdxSanity
{
    private static readonly XNamespace ct  = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace pr  = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace v   = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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

        // [Content_Types].xml
        var ctDoc = Load(pkg, "/[Content_Types].xml");
        var defaults  = ctDoc.Root!.Elements(ct + "Default").ToList();
        var overrides = ctDoc.Root!.Elements(ct + "Override").ToList();

        var xmlDefault = defaults.FirstOrDefault(d => (string)d.Attribute("Extension") == "xml");
        if (xmlDefault == null || (string)xmlDefault.Attribute("ContentType") != "application/xml")
            issues.Add("Default .xml must be application/xml; use Overrides for Visio parts.");

        bool HasOverride(string partName, string contentType) =>
            overrides.Any(o => (string)o.Attribute("PartName") == partName &&
                               (string)o.Attribute("ContentType") == contentType);

        if (!HasOverride("/visio/document.xml", CT_Document))
            issues.Add("Missing Override for /visio/document.xml -> application/vnd.ms-visio.drawing.main+xml.");
        if (!HasOverride("/visio/pages/pages.xml", CT_Pages))
            issues.Add("Missing Override for /visio/pages/pages.xml -> application/vnd.ms-visio.pages+xml.");
        if (!HasOverride("/visio/pages/page1.xml", CT_Page))
            issues.Add("Missing Override for /visio/pages/page1.xml -> application/vnd.ms-visio.page+xml.");

        // Root rels
        var rootRels = Load(pkg, "/_rels/.rels");
        var docRel = rootRels.Root!.Elements(pr + "Relationship")
            .FirstOrDefault(r => (string)r.Attribute("Target") == "/visio/document.xml"
                              || (string)r.Attribute("Target") == "visio/document.xml");
        if (docRel == null || (string)docRel.Attribute("Type") != RT_Document)
            issues.Add("Package → /visio/document.xml must use Visio document rel type.");

        // Document → Pages
        var docRels = Load(pkg, "/visio/_rels/document.xml.rels");
        var pagesRel = docRels.Root!.Elements(pr + "Relationship")
            .FirstOrDefault(r => (string)r.Attribute("Target") == "pages/pages.xml"
                              || (string)r.Attribute("Target") == "/visio/pages/pages.xml");
        if (pagesRel == null || (string)pagesRel.Attribute("Type") != RT_Pages)
            issues.Add("document.xml → pages/pages.xml must use Visio pages rel type.");

        // Pages → Page1
        var pagesRels = Load(pkg, "/visio/pages/_rels/pages.xml.rels");
        var pageRel = pagesRels.Root!.Elements(pr + "Relationship")
            .FirstOrDefault(r => ((string)r.Attribute("Target"))?.EndsWith("page1.xml", StringComparison.OrdinalIgnoreCase) == true);
        if (pageRel == null || (string)pageRel.Attribute("Type") != RT_Page)
            issues.Add("pages.xml → page1.xml must use Visio page rel type.");

        // Pages XML structure
        var pagesXml = Load(pkg, "/visio/pages/pages.xml");
        var page = pagesXml.Root!.Element(v + "Page");
        if (page == null) issues.Add("pages.xml must contain a <Page> element.");
        else {
            if (!int.TryParse((string)page.Attribute("ID"), out var pid) || pid < 1)
                issues.Add("Page/@ID must be numeric and 1-based.");
            var relChild = page.Element(v + "Rel");
            var rid = (string?)relChild?.Attribute(rel + "id");
            if (rid == null || !rid.StartsWith("rId", StringComparison.Ordinal))
                issues.Add("Page must contain child <Rel r:id=\"rId#\"/> (not an attribute).");
        }

        // Page: all Shape IDs numeric
        var page1 = Load(pkg, "/visio/pages/page1.xml");
        var badShape = page1.Descendants(v + "Shape")
            .Select(s => (string)s.Attribute("ID"))
            .FirstOrDefault(id => !int.TryParse(id, out _));
        if (badShape != null) issues.Add($"Shape/@ID must be numeric. Found '{badShape}'.");

        return issues;
    }

    private static XDocument Load(Package pkg, string partName)
    {
        using var s = pkg.GetPart(new Uri(partName, UriKind.Relative)).GetStream();
        return XDocument.Load(s);
    }
}
```

---

## What to change in your repo (summary checklist)

* [ ] **Remove** any code that writes `[Content_Types].xml` (or changes Default for `.xml`).
  If you keep it, ensure:
  **Default `.xml` → `application/xml`**, plus **Overrides** for
  `/visio/document.xml`, `/visio/pages/pages.xml`, `/visio/pages/page1.xml`. ([c-rex.net][3])
* [ ] Keep the **relationship types** exactly as you have them now (they’re correct). ([Microsoft Learn][1])
* [ ] Use **relative** `Target` URIs from each part (`pages/pages.xml`, `page1.xml`). ([Wikipedia][4])
* [ ] Ensure every `Shape/@ID` is **integer**. No `C1`, `S2`, etc.
* [ ] `pages.xml` must use `<Rel r:id="rId#"/>` inside `<Page>`. (You already do this.) ([Microsoft Learn][2])

---

## Why this works (sources)

* Document part and its relationship type and content type: Microsoft MS‑VSDX spec. ([Microsoft Learn][1])
* Pages and Page parts, relationships, and required **PageContents** root: Microsoft MS‑VSDX spec. ([Microsoft Learn][2])
* Root element for `document.xml` is **VisioDocument** in `http://schemas.microsoft.com/office/visio/2012/main`. ([Microsoft Learn][5])
* OPC content types model (**Default** for extensions, **Override** for part names): OOXML/OPC docs. ([c-rex.net][3])
* Visio VSDX is an OPC package; follow OPC rules for relationships/targets. ([Microsoft Learn][6])

---

## What I’d do next (straight to the point)

1. **Delete** your custom `[Content_Types].xml` writer.
2. Ensure you create the three parts with the exact content types; **let OPC emit overrides**.
3. Keep your current relationships; change `Target` values to **relative**.
4. Ensure **numeric** shape IDs (fix “C1” in your connector code path).
5. Drop the `VsdxSanity.Validate(path)` test in your CI; fail the build on any issues.

Do those five steps; Visio will open the files.

[1]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/7ec3d7b0-0de2-4711-a7b6-92daa2020d71?utm_source=chatgpt.com "Document XML Part - MS-VSDX"
[2]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/1f15c8f0-6565-465c-aefd-2be6af545e8a?utm_source=chatgpt.com "[MS-VSDX]: Page XML Part"
[3]: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Content_topic_ID0ESTAI.html?utm_source=chatgpt.com "Content Types Stream Markup"
[4]: https://en.wikipedia.org/wiki/Open_Packaging_Conventions?utm_source=chatgpt.com "Open Packaging Conventions"
[5]: https://learn.microsoft.com/en-us/office/client-developer/visio/visiodocument-elementvisio-xml?utm_source=chatgpt.com "VisioDocument element (Visio XML)"
[6]: https://learn.microsoft.com/en-us/office/client-developer/visio/introduction-to-the-visio-file-formatvsdx?utm_source=chatgpt.com "Introduction to the Visio file format (.vsdx)"
