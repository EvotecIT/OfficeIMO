"""Independent pypdf readback for the tagged deterministic six/20-column PDF workloads.

Usage: python verify-pdf-workload-structure.py <pdf> <source-rows> <columns> [--embedded-fonts]
Physical every-cell/band checks remain with the canonical Poppler workload verifier.
"""
from collections import Counter
import hashlib
import importlib.util
import json
from pathlib import Path
import sys

from pypdf import PdfReader
from pypdf.generic import DictionaryObject

spec = importlib.util.spec_from_file_location("semantic_report", Path(__file__).with_name("verify-semantic-report.py"))
semantic = importlib.util.module_from_spec(spec)
spec.loader.exec_module(semantic)
require, resolved, reference, sequence = semantic.require, semantic.resolved, semantic.reference, semantic.sequence

HEADERS = ["ID", "Asset", "Owner", "Status", "Amount", "Seen", "Protocol", "Version", "Region", "CPU", "Memory", "Storage", "Requests", "Errors", "TLS", "Backup", "Ticket", "Policy", "Zone", "Health"]


def expected_row(index, columns):
    values = ["N" + str(index).zfill(6), "node-" + str(index).zfill(6), "ops-" + str(index % 17),
              ["Pass", "Warning", "Critical"][index % 3], str(index + .5), "2026-01-" + str(index % 28 + 1).zfill(2)]
    return values + ["M" + str(column) + "-" + str(index).zfill(6) for column in range(6, columns)]


def verify_embedded_fonts(reader):
    fonts = {}
    for page in reader.pages:
        resources = resolved(page["/Resources"])
        page_fonts = resolved(resources.get("/Font"))
        require(isinstance(page_fonts, DictionaryObject) and bool(page_fonts), "Embedded-font workload has no page fonts")
        for value in page_fonts.values():
            identity = reference(value)
            require(identity is not None, "Embedded font must have an indirect resource")
            if identity in fonts:
                continue
            font = resolved(value)
            require(str(font.get("/Subtype")) == "/Type0" and str(font.get("/Encoding")) == "/Identity-H", "Embedded workload used a standard font instead of a Type0 Identity-H face")
            unicode_map = resolved(font.get("/ToUnicode"))
            require(hasattr(unicode_map, "get_data") and b"begincmap" in unicode_map.get_data(), "Embedded font has no ToUnicode map")
            descendants = sequence(font.get("/DescendantFonts"))
            require(len(descendants) == 1, "Embedded font must have one CID descendant")
            descendant = resolved(descendants[0])
            require(str(descendant.get("/Subtype")) == "/CIDFontType2", "Embedded font has no TrueType CID descendant")
            descriptor = resolved(descendant.get("/FontDescriptor"))
            require(isinstance(descriptor, DictionaryObject), "Embedded font has no descriptor")
            stream = resolved(descriptor.get("/FontFile2"))
            require(hasattr(stream, "get_data"), "Embedded font has no FontFile2 payload")
            payload = stream.get_data()
            require(bool(payload) and int(stream.get("/Length1", -1)) == len(payload), "Embedded font payload length differs")
            fonts[identity] = {"baseFont": str(font.get("/BaseFont")), "decodedBytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest()}
    require(bool(fonts), "Embedded workload produced no fonts")
    return {"type0IdentityH": True, "toUnicode": True, "cidFontType2": True, "fontFile2": True,
            "resources": list(fonts.values()), "fontPayloadBoundary": "Subset font payload hashes are recorded; equality with the full source fonts is not claimed."}


def verify_reader(reader, rows, columns, embedded_fonts=False):
    require(rows >= 0 and columns in (6, 20), "Declare the deterministic six/20-column workload")
    catalog = reader.trailer["/Root"]
    require(bool(catalog["/MarkInfo"]["/Marked"]) and str(catalog["/Lang"]) == "en-US", "Tagged inventory language/marked contract differs")
    root = catalog["/StructTreeRoot"]
    parents = dict((int(key), value) for key, value in semantic.tree_entries(root["/ParentTree"], "/Nums"))
    identifiers = dict((str(key), value) for key, value in semantic.tree_entries(root["/IDTree"], "/Names"))
    pages = {reference(page): page for page in reader.pages}
    marks = {identity: semantic.page_marks(page, reader) for identity, page in pages.items()}
    nodes, consumed = {}, set()

    def mark(page_key, mcid, parent):
        require(page_key in pages and mcid in marks[page_key], "Logical mark has no physical content")
        key = page_key, mcid
        require(key not in consumed, "A physical mark has duplicate logical owners")
        consumed.add(key)
        owners = sequence(parents[int(pages[page_key]["/StructParents"])])
        require(mcid < len(owners) and reference(owners[mcid]) == parent, "Marked-content ParentTree owner differs")
        return marks[page_key][mcid]

    def visit(value, parent, inherited_page=None):
        obj = resolved(value)
        if isinstance(obj, int):
            return mark(inherited_page, obj, parent)
        require(isinstance(obj, DictionaryObject), "Unexpected logical child value")
        page_key = reference(obj.get("/Pg")) or inherited_page
        if str(obj.get("/Type")) == "/MCR":
            return mark(page_key, int(obj["/MCID"]), parent)
        if "/S" not in obj:
            return "".join(visit(child, parent, page_key) for child in sequence(obj.get("/K")))
        identity = reference(value)
        require(identity is not None and identity not in nodes, "Duplicate/non-indirect structure element")
        require(reference(obj["/P"]) == parent, "Logical structure parent differs")
        entry = {"object": obj, "role": str(obj["/S"]), "children": sequence(obj.get("/K"))}
        nodes[identity] = entry
        entry["text"] = "".join(visit(child, identity, page_key) for child in entry["children"])
        return entry["text"]

    visit(root, reference(root))
    require(len(consumed) == sum(len(value) for value in marks.values()), "Marked physical content is absent from the logical tree")
    roles = Counter(entry["role"] for entry in nodes.values())
    require(roles == Counter({"/Document": 1, "/Table": 1, "/TR": rows + 1, "/TH": columns, "/TD": rows * columns}), "Logical table role/row/cell counts differ")
    logical_rows = [entry for entry in nodes.values() if entry["role"] == "/TR"]
    header_cells = [nodes[reference(child)] for child in logical_rows[0]["children"]]
    require([entry["text"] for entry in header_cells] == HEADERS[:columns], "Logical column headings/order differ")
    require(all(str(semantic.attributes(entry["object"]).get("/Scope")) == "/Column" for entry in header_cells), "Column heading scope differs")
    for name, target in identifiers.items():
        identity = reference(target)
        require(identity in nodes and str(nodes[identity]["object"].get("/ID")) == name and nodes[identity]["role"] == "/TH", "Header identity target differs")
    header_owners = [reference(entry["object"]) for entry in header_cells]
    digest = hashlib.sha256()
    digest.update((json.dumps(HEADERS[:columns], separators=(",", ":")) + "\n").encode())
    for index, row in enumerate(logical_rows[1:]):
        cells = [nodes[reference(child)] for child in row["children"]]
        actual = [entry["text"] for entry in cells]
        require(actual == expected_row(index, columns), "Logical cell value/order differs at row " + str(index))
        for column, cell in enumerate(cells):
            headers = sequence(semantic.attributes(cell["object"]).get("/Headers"))
            require(len(headers) == 1 and str(headers[0]) in identifiers and reference(identifiers[str(headers[0])]) == header_owners[column], "Body cell header association differs")
        digest.update((json.dumps(actual, separators=(",", ":")) + "\n").encode())
    return {"dataCells": rows * columns, "logicalRows": rows, "columnHeaders": columns, "logicalElements": len(nodes), "roles": dict(roles),
            "markedFragments": len(consumed), "logicalSha256": digest.hexdigest(), "logicalValuesAndOrder": True,
            "headerAssociations": True, "uniqueMarkedOwners": True, "parentTree": True, "independentReader": "pypdf",
            **({"embeddedFonts": verify_embedded_fonts(reader)} if embedded_fonts else {})}


if __name__ == "__main__":
    require(sys.argv[4:] in ([], ["--embedded-fonts"]), "Unknown workload verification option")
    print(json.dumps(verify_reader(PdfReader(sys.argv[1], strict=True), int(sys.argv[2]), int(sys.argv[3]), embedded_fonts=bool(sys.argv[4:]))))
