#!/usr/bin/env python3
"""Capture the same fixture and test files at two revisions, then compare them.

Use a fresh evidence directory for each capture. Run with the same SDK, time
zone and fixture set on both revisions. The runner may live in either checkout.

    python3 Build/Compare-LegacyXls.py capture /path/to/master /scratch/master
    python3 Build/Compare-LegacyXls.py capture /path/to/candidate /scratch/candidate
    python3 Build/Compare-LegacyXls.py compare /scratch/master /scratch/candidate

XLS bytes are compared directly, apart from the generated OLE save time.
For Open XML, compare each part after normalizing temporary relationship IDs,
generated core metadata, ZIP and XML serialization, the order of unique OPC
content type declarations, and namespace prefixes in plain worksheets.
The encryption test compares the decrypted package and retains the ciphertext.
"""

import argparse
import hashlib
import json
import os
import re
import struct
from pathlib import Path, PurePosixPath
import subprocess
import sys
from xml.dom import minidom
import xml.etree.ElementTree as ET
import zipfile

PROBE = Path(__file__).resolve().with_name("LegacyXlsProof") / "Capture.cs"
TEST_PROJECT = "OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj"
TEST_CLASS = "OfficeIMO.Tests.Excel.LegacyXls"
DELETE_HELPER = (
    "        private static void TryDelete(string path) {\n"
    "            if (File.Exists(path)) {\n"
    "                File.Delete(path);\n"
    "            }\n"
    "        }"
)


def capture(checkout, output, build_lock):
    if build_lock:
        subprocess.run([
            "lockf", "-k", str(build_lock), sys.executable, str(Path(__file__).resolve()),
            "capture", str(checkout), str(output),
        ], check=True)
        return
    output.mkdir(parents=True)
    capture_source = checkout / "OfficeIMO.Excel.Tests/Excel.LegacyXls.NormalLoad.cs"
    source = capture_source.read_text(encoding="utf-8")
    if source.count(DELETE_HELPER) != 1:
        raise RuntimeError("The existing LegacyXls cleanup helper changed")
    instrumented = (
        "        private static void TryDelete(string path,\n"
        "            [System.Runtime.CompilerServices.CallerMemberName] string test = \"\") =>\n"
        "            LegacyXlsProof.Delete(path, test);"
    )
    hooks = output / "capture.targets"
    shadow = output / capture_source.name
    shadow.write_text(source.replace(DELETE_HELPER, instrumented), encoding="utf-8")
    targets = ET.Element("Project")
    group = ET.SubElement(targets, "ItemGroup", Condition="'$(MSBuildProjectName)' == 'OfficeIMO.Excel.Tests'")
    ET.SubElement(group, "Compile", Remove=capture_source.name)
    ET.SubElement(group, "Compile", Include=str(shadow))
    ET.SubElement(group, "Compile", Include=str(PROBE))
    ET.ElementTree(targets).write(hooks, encoding="utf-8")
    revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=checkout, text=True).strip()
    env = dict(os.environ, OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT=str(output),
               OFFICEIMO_TEST_REPOSITORY_ROOT=str(checkout), TZ="UTC")
    subprocess.run([
        "dotnet", "build", TEST_PROJECT, "-c", "Release", "-f", "net10.0", "-m:2",
        "-p:CustomAfterMicrosoftCommonTargets=" + str(hooks),
    ], cwd=checkout, env=env, check=True)
    subprocess.run([
        "dotnet", "test", TEST_PROJECT, "-c", "Release", "-f", "net10.0", "--no-build", "-m:2",
        "--filter", (
            f"(FullyQualifiedName~{TEST_CLASS}|FullyQualifiedName~ExcelOpenXmlSdkBoundary"
            "|FullyQualifiedName~OfficeIMO.Tests.Excel.WorkbookWriter_Xls"
            "|FullyQualifiedName~OfficeIMO.Tests.Excel.OpenDataReader_XlsSelectedSheetMatchesTheSameSheetReadWithTheWholeWorkbook"
            ")&Category!=Performance"
        ),
        "--logger", "trx;LogFileName=proof.trx", "--results-directory", str(output),
    ], cwd=checkout, env=env, check=True)
    encrypted = output / "test-files" / ENCRYPTED_PACKAGE_TEST / "001.xlsx"
    if encrypted.is_file():
        subprocess.run([
            "pwsh", "-NoLogo", "-NoProfile", "-File", str(PROBE.with_name("Decrypt-Fixture.ps1")),
            "-CoreAssembly", str(checkout / "OfficeIMO.Core/bin/Release/net10.0/OfficeIMO.Core.dll"),
            "-Package", str(encrypted),
        ], check=True)
    if revision != subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=checkout, text=True).strip():
        raise RuntimeError("The checkout revision changed during capture")
    (output / "revision.txt").write_text(revision + "\n", encoding="utf-8")
    if not any((output / "fixtures").glob("**/*.xlsx")) or not any((output / "test-files").glob("**/*.xls")):
        raise RuntimeError("The probe did not capture both fixture projections and test-saved XLS files")


def sha(data):
    return hashlib.sha256(data).hexdigest()


COMPOUND_HEADER = bytes.fromhex("d0cf11e0a1b11ae1")
SUMMARY_INFORMATION = bytes.fromhex("e0859ff2f94f6810ab9108002b27b3d9")
CORE_PROPERTIES_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
GENERATED_CORE_PROPERTIES = re.compile(r"package/services/metadata/core-properties/[0-9a-f]{32}\.psmdcp")
STABLE_CORE_PROPERTIES = "package/services/metadata/core-properties/core.psmdcp"
ENCRYPTED_PACKAGE_TEST = "LegacyXls_NormalLoad_SaveEncryptedPathProducesEncryptedXlsx"
MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XMLNS = "http://www.w3.org/2000/xmlns/"
XML_SCHEMA_INSTANCE = "http://www.w3.org/2001/XMLSchema-instance"
XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"
SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
RELATIONSHIP_NAMESPACES = frozenset({
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
})


def relationship_owner(name):
    part = PurePosixPath(name)
    return part.parent.parent / part.name[:-5]


def canonical_relationships(data, aliases):
    dom = minidom.parseString(data)
    root = dom.documentElement
    relationships = root.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS, "Relationship")
    relationships = [item for item in relationships if item.parentNode is root]

    def target(item):
        return tuple(item.getAttribute(name) for name in ("Type", "Target", "TargetMode"))

    for item in relationships:
        if item.getAttribute("Type") == CORE_PROPERTIES_REL:
            original = item.getAttribute("Target")
            normalized = aliases.get(original.lstrip("/"))
            if normalized is not None:
                item.setAttribute("Target", ("/" if original.startswith("/") else "") + normalized)

    canonical_ids = {item: f"rId{index}" for index, item in enumerate(
        sorted({target(item) for item in relationships}), start=1)}
    mapping = {}
    for item in sorted(relationships, key=target):
        original = item.getAttribute("Id")
        if not original or original in mapping:
            raise ValueError("Missing or duplicate relationship ID")
        mapping[original] = canonical_ids[target(item)]
        item.setAttribute("Id", mapping[original])
        root.appendChild(item)
    return mapping, dom.toxml(encoding="utf-8")


def normalize_xml(data, ids, generated_save_time, aliases=None, generated_creation_time=False):
    if not ids and not generated_save_time and not aliases:
        return data
    dom = minidom.parseString(data)
    if ids:
        for element in dom.getElementsByTagName("*"):
            for attribute in element.attributes.values():
                if attribute.namespaceURI in RELATIONSHIP_NAMESPACES or (
                    attribute.namespaceURI == "urn:schemas-microsoft-com:office:office"
                    and attribute.localName == "relid"
                ):
                    attribute.value = ids.get(attribute.value, "unresolved:" + attribute.value)
    if aliases:
        for item in dom.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/content-types", "Override"):
            name = item.getAttribute("PartName").lstrip("/")
            if name in aliases:
                item.setAttribute("PartName", "/" + aliases[name])
    if generated_save_time:
        # Compare authored creation except in the test that sets it from UtcNow.
        terms = ("modified", "created") if generated_creation_time else ("modified",)
        for term in terms:
            for element in dom.getElementsByTagNameNS("http://purl.org/dc/terms/", term):
                if len(element.childNodes) == 1 and element.firstChild.nodeType == element.TEXT_NODE:
                    element.firstChild.data = "(generated time)"
    return dom.toxml(encoding="utf-8")


def namespace_value_bindings(data):
    # Canonical XML omits namespace declarations that appear only in QName
    # or prefix-valued attributes. Those bindings can change markup selection.
    dom = minidom.parseString(data)
    bindings = []
    prefix_lists = {"Ignorable", "MustUnderstand"}
    qname_lists = {"ProcessContent", "PreserveElements", "PreserveAttributes"}
    for index, element in enumerate(dom.getElementsByTagName("*")):
        compatibility_attributes = [(name, element.getAttributeNS(MARKUP_COMPATIBILITY, name))
                                    for name in prefix_lists | qname_lists]
        if element.namespaceURI == MARKUP_COMPATIBILITY and element.localName == "Choice":
            compatibility_attributes.append(("Requires", element.getAttribute("Requires")))
        if element.hasAttributeNS(XML_SCHEMA_INSTANCE, "type"):
            compatibility_attributes.append(("xsi:type", element.getAttributeNS(XML_SCHEMA_INSTANCE, "type")))
        for name, value in sorted(compatibility_attributes):
            for token in value.split():
                if name in qname_lists or name == "xsi:type":
                    prefix, separator, _ = token.partition(":")
                    if not separator:
                        prefix = ""
                else:
                    prefix = token
                scope = element
                uri = XML_NAMESPACE if prefix == "xml" else None
                while scope is not None and scope.nodeType == scope.ELEMENT_NODE:
                    if scope.hasAttributeNS(XMLNS, prefix or "xmlns"):
                        uri = scope.getAttributeNS(XMLNS, prefix or "xmlns")
                        break
                    scope = scope.parentNode
                bindings.append((index, name, token, uri))
    return bindings


PLAIN_WORKSHEET_ATTRIBUTES = {
    "worksheet": (),
    "dimension": ("ref",),
    "sheetViews": (),
    "sheetView": ("workbookViewId",),
    "selection": ("activeCell", "sqref"),
    "sheetData": (),
    "row": ("r",),
    "c": ("r", "t", "s"),
    "v": (),
}


def canonical_content_types(data):
    # OPC defaults/overrides are keyed declarations. Their XML order has no
    # meaning when the keys are unique; anything else uses the strict comparison.
    dom = minidom.parseString(data)
    root = dom.documentElement
    if (root.namespaceURI, root.localName) != (CONTENT_TYPES, "Types"):
        return data
    declarations = {}
    for item in root.childNodes:
        if item.nodeType == item.TEXT_NODE and not item.data.strip():
            continue
        if item.nodeType != item.ELEMENT_NODE or item.namespaceURI != CONTENT_TYPES or item.hasChildNodes():
            return data
        key_attribute = {"Default": "Extension", "Override": "PartName"}.get(item.localName)
        if key_attribute is None or set(item.attributes.keys()) != {key_attribute, "ContentType"}:
            return data
        key = (item.localName, item.getAttribute(key_attribute).casefold())
        if key in declarations:
            return data
        declarations[key] = item
    for key in sorted(declarations):
        root.appendChild(declarations[key])
    return dom.toxml(encoding="utf-8")


def known_plain_worksheet(data):
    # The primitive projection only produces this vocabulary. QName-valued or
    # compatibility attributes and extension elements take the strict comparison.
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True))
    root = ET.fromstring(data, parser=parser)
    prefix = "{" + SPREADSHEET + "}"
    if root.tag != prefix + "worksheet":
        return False
    for node in root.iter():
        if not isinstance(node.tag, str) or not node.tag.startswith(prefix):
            return False
        attributes = PLAIN_WORKSHEET_ATTRIBUTES.get(node.tag[len(prefix):])
        if attributes is None or any(name not in attributes for name in node.attrib):
            return False
    return True


def canonical_xml_hash(data, name=None):
    if name == "[Content_Types].xml":
        data = canonical_content_types(data)
    plain_sheet = bool(name and name.startswith("xl/worksheets/") and name.endswith(".xml")
                       and known_plain_worksheet(data))
    payload = ET.canonicalize(data, with_comments=True, rewrite_prefixes=plain_sheet).encode("utf-8")
    bindings = namespace_value_bindings(data)
    if bindings:
        payload += b"\0" + json.dumps(bindings, separators=(",", ":")).encode("utf-8")
    return sha(payload)


def package_manifest(path):
    manifest = {}
    with zipfile.ZipFile(path) as package:
        members = package.infolist()
        if len({item.filename for item in members}) != len(members):
            raise ValueError(f"Duplicate package part in {path}")
        aliases = {}
        names = {item.filename for item in members}
        if "_rels/.rels" in names:
            root = minidom.parseString(package.read("_rels/.rels"))
            core = [item for item in root.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS, "Relationship")
                    if item.getAttribute("Type") == CORE_PROPERTIES_REL]
            if len(core) == 1:
                name = core[0].getAttribute("Target").lstrip("/")
                if GENERATED_CORE_PROPERTIES.fullmatch(name) and name in names and STABLE_CORE_PROPERTIES not in names:
                    aliases[name] = STABLE_CORE_PROPERTIES

        relationships = {}
        for member in members:
            name = member.filename
            if PurePosixPath(name).parent.name == "_rels" and name.endswith(".rels"):
                ids, canonical = canonical_relationships(package.read(member), aliases if name == "_rels/.rels" else {})
                relationships[relationship_owner(name)] = ids
                manifest[name] = canonical_xml_hash(canonical)

        for member in members:
            name = member.filename
            if name in manifest:
                continue
            data = package.read(member)
            is_core_properties = name == "docProps/core.xml" or name in aliases
            canonical_name = aliases.get(name, name)
            if name.endswith((".xml", ".vml")) or is_core_properties:
                # SDK-backed saves generate different identifiers for the same part targets.
                # Resolve references through their own part's relationship set.
                ids = relationships.get(PurePosixPath(name), {})
                generated_creation_time = path.parent.name == "LegacyXls_DocumentProperties_ProjectScalarCustomPropertyVariantTypes"
                data = normalize_xml(data, ids, is_core_properties,
                                     aliases if name == "[Content_Types].xml" else None,
                                     generated_creation_time)
                # Preserve sibling order, text, attributes, comments and unknown
                # markup. Ignore ZIP layout, XML attribute order and declarations.
                manifest[canonical_name] = canonical_xml_hash(data, name)
            elif (name == "OfficeIMOCompatibility/SourceMetadata"
                  and path.parent.name == "LegacyXls_Convert_PreservationOnlyRetainsOriginalInXlsxCarrier"):
                manifest[name] = sha(normalize_source_metadata(data))
            else:
                manifest[name] = sha(data)
    return manifest


def normalize_source_metadata(data):
    # The preservation test writes its source to a new GUID-named temporary file.
    # Keep the format, payload checksum, mode and extension in the comparison.
    prefix = b"\x1cOfficeIMOCompatibilitySource\x01\x00\x00\x00\x09Excel.Xls\x24"
    if data.startswith(prefix) and re.match(rb"[0-9a-f]{32}\.xls@", data[len(prefix):]):
        return data[:len(prefix)] + b"0" * 32 + data[len(prefix) + 32:]
    return data


def normalize_xls_save_time(path):
    data = path.read_bytes()
    if not data.startswith(COMPOUND_HEADER):
        return data
    # Look for the single-section, little-endian SummaryInformation property set.
    # BIFF records, other OLE streams and metadata remain byte-for-byte evidence.
    # A FILETIME must fall near the captured write to qualify as generated.
    filetime_epoch = 11644473600
    written = path.stat().st_mtime
    positions = []
    search_from = 28
    while (fmtid := data.find(SUMMARY_INFORMATION, search_from)) >= 0:
        search_from = fmtid + len(SUMMARY_INFORMATION)
        header = fmtid - 28
        if data[header:header + 2] != b"\xfe\xff" or data[fmtid - 4:fmtid] != b"\x01\x00\x00\x00":
            continue
        if fmtid + 20 > len(data):
            continue
        section = header + struct.unpack_from("<I", data, fmtid + 16)[0]
        if not fmtid + 20 <= section <= len(data) - 8:
            continue
        size, count = struct.unpack_from("<II", data, section)
        if section + size > len(data) or 8 + count * 8 > size:
            continue
        for entry in range(section + 8, section + 8 + count * 8, 8):
            kind, offset = struct.unpack_from("<II", data, entry)
            if kind != 13 or not 8 + count * 8 <= offset <= size - 12:
                continue
            value_type, ticks = struct.unpack_from("<IQ", data, section + offset)
            if value_type == 64 and abs((ticks / 10_000_000 - filetime_epoch) - written) < 600:
                positions.append(section + offset + 4)
    if len(positions) == 1:
        offset = positions[0]
        return data[:offset] + b"\0" * 8 + data[offset + 8:]
    return data


def artifacts(directory):
    return {
        path.relative_to(directory).as_posix(): path
        for folder in ("fixtures", "test-files")
        for path in (directory / folder).rglob("*") if path.is_file()
    }


def compare(master, candidate):
    before, after = artifacts(master), artifacts(candidate)
    if not before or not after:
        raise RuntimeError("Both captures must contain evidence")
    identical, packages, compound, encrypted, differences = 0, 0, 0, 0, {}
    for name in sorted(before.keys() | after.keys()):
        if name not in before or name not in after:
            differences[name] = "missing from " + ("master" if name not in before else "candidate")
        elif before[name].read_bytes() == after[name].read_bytes():
            identical += 1
        elif zipfile.is_zipfile(before[name]) and zipfile.is_zipfile(after[name]):
            first, second = package_manifest(before[name]), package_manifest(after[name])
            changed = sorted(part for part in first.keys() | second.keys() if first.get(part) != second.get(part))
            if changed:
                differences[name] = changed
            else:
                packages += 1
        elif name == f"test-files/{ENCRYPTED_PACKAGE_TEST}/001.xlsx":
            witness = name.removesuffix(".xlsx") + ".plaintext.xlsx"
            if (witness in before and witness in after
                    and package_manifest(before[witness]) == package_manifest(after[witness])):
                encrypted += 1
            else:
                differences[name] = "randomized ciphertext requires matching plaintext evidence"
        elif Path(name).suffix == ".xls" and normalize_xls_save_time(before[name]) == normalize_xls_save_time(after[name]):
            compound += 1
        else:
            differences[name] = "different bytes"
    print(json.dumps({"byte_identical": identical, "equivalent_packages": packages,
                      "equivalent_compound_files": compound, "equivalent_encrypted_packages": encrypted,
                      "differences": differences}, indent=2))
    return bool(differences)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(required=True, dest="command")
    generate = commands.add_parser("capture")
    generate.add_argument("checkout", type=Path)
    generate.add_argument("output", type=Path)
    generate.add_argument("--build-lock", type=Path, help="Shared lockf path for the local build and no-build test")
    diff = commands.add_parser("compare")
    diff.add_argument("master", type=Path)
    diff.add_argument("candidate", type=Path)
    args = parser.parse_args()
    if args.command == "capture":
        capture(args.checkout.resolve(), args.output.resolve(), args.build_lock.expanduser().resolve() if args.build_lock else None)
        return 0
    return int(compare(args.master, args.candidate))


if __name__ == "__main__":
    raise SystemExit(main())
