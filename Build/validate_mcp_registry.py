"""Validate OfficeIMO's Registry metadata and NuGet ownership contract."""

import argparse
import json
from pathlib import Path
from urllib.request import urlopen
from xml.etree import ElementTree
from zipfile import ZipFile


SCHEMA_URL = "https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json"
ROOT = Path(__file__).resolve().parent.parent


def validate_metadata(server, schema):
    from jsonschema import Draft7Validator, FormatChecker

    if server.get("$schema") != SCHEMA_URL:
        raise ValueError("Registry metadata must use the pinned official schema")
    Draft7Validator(schema, format_checker=FormatChecker()).validate(server)


def validate_release_versions(server, plugin_version):
    packages = [package for package in server.get("packages", [])
                if package.get("registryType") == "nuget" and package.get("identifier") == "OfficeIMO.Tool"]
    if len(packages) != 1:
        raise ValueError("Registry metadata must declare exactly one OfficeIMO.Tool NuGet package")
    if server.get("version") != plugin_version or packages[0].get("version") != plugin_version:
        raise ValueError("Registry, NuGet package and plugin versions must match")


def validate_marker(readme, server):
    marker = f'<!-- mcp-name: {server["name"]} -->'
    markers = [line.strip() for line in readme.splitlines() if "mcp-name:" in line]
    if markers != [marker]:
        raise ValueError(f"README must contain exactly one ownership marker: {marker}")


def validate_package(package, server):
    with ZipFile(package) as archive:
        nuspecs = [name for name in archive.namelist() if name.endswith(".nuspec")]
        if len(nuspecs) != 1:
            raise ValueError("Expected one NuGet package manifest")
        manifest = ElementTree.fromstring(archive.read(nuspecs[0]))
        metadata = manifest.find("{*}metadata")
        if metadata is None or metadata.findtext("{*}id") != "OfficeIMO.Tool":
            raise ValueError("Expected the OfficeIMO.Tool package")
        readme_path = metadata.findtext("{*}readme")
        if not readme_path:
            raise ValueError("The NuGet package must declare a README")
        validate_marker(archive.read(readme_path).decode("utf-8-sig"), server)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("metadata", "package"))
    parser.add_argument("package", nargs="?")
    args = parser.parse_args()
    server = json.loads((ROOT / "OfficeIMO.Tool/server.json").read_text())
    if args.mode == "package":
        if not args.package:
            parser.error("package mode requires a .nupkg path")
        validate_package(args.package, server)
    else:
        if server.get("$schema") != SCHEMA_URL:
            raise ValueError("Registry metadata must use the pinned official schema")
        with urlopen(SCHEMA_URL, timeout=30) as response:
            schema = json.load(response)
        validate_metadata(server, schema)
        plugin = json.loads((ROOT / ".agents/plugins/officeimo-document-tools/plugin.json").read_text())
        validate_release_versions(server, plugin["version"])
        validate_marker((ROOT / "OfficeIMO.Tool/README.md").read_text(), server)


if __name__ == "__main__":
    main()
