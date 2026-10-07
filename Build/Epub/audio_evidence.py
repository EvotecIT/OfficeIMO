"""Optional, bounded encoded-audio evidence for the EPUB validation runner."""
from decimal import Decimal, InvalidOperation, localcontext
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import subprocess
import tempfile
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET
import zipfile

SMIL = "{http://www.w3.org/ns/SMIL}"
OPF = "{http://www.idpf.org/2007/opf}"
CONTAINER = "{urn:oasis:names:tc:opendocument:xmlns:container}"
DEMUXERS = {"audio/mpeg": "mp3", "audio/mp4": "mov", "audio/wav": "wav", "audio/x-wav": "wav", "audio/ogg": "ogg", "audio/flac": "flac"}


class UncheckedAudio(ValueError):
    """The requested check cannot establish this input's encoded-audio bounds."""


def clock(value):
    with localcontext() as context:
        context.prec = 160
        return _clock(value)


def _clock(value):
    value = value.strip()
    if len(value) > 128:
        raise ValueError("SMIL clock exceeds its length bound")
    if re.fullmatch(r"[0-9]+:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?", value):
        h, m, s = map(Decimal, value.split(":"))
        if m >= 60 or s >= 60:
            raise ValueError("Invalid SMIL clock components")
        return h * 3600 + m * 60 + s
    if re.fullmatch(r"[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?", value):
        m, s = map(Decimal, value.split(":"))
        if m >= 60 or s >= 60:
            raise ValueError("Invalid SMIL clock components")
        return m * 60 + s
    match = re.fullmatch(r"([0-9]+(?:\.[0-9]+)?)(h|min|ms|s)?", value)
    if not match:
        raise ValueError("Unsupported or invalid SMIL clock")
    return Decimal(match[1]) * {None: Decimal(1), "s": Decimal(1), "ms": Decimal(".001"), "min": Decimal(60), "h": Decimal(3600)}[match[2]]


def local_path(base, reference):
    if not reference or "\\" in reference or re.search(r"%(?:2f|5c)", reference, re.I):
        raise ValueError("Ambiguous audio resource path")
    parsed = urlsplit(reference)
    if parsed.scheme or parsed.netloc or parsed.path.startswith("/") or parsed.query or parsed.fragment:
        raise UncheckedAudio("External, rooted, query or fragment audio references are not checked")
    decoded = unquote(parsed.path, errors="strict")
    if "\x00" in decoded or "\\" in decoded:
        raise ValueError("Invalid audio resource path")
    parts = list(PurePosixPath(base).parent.parts) if base else []
    for part in decoded.split("/"):
        if part == "..":
            if not parts:
                raise ValueError("Audio resource escapes the container")
            parts.pop()
        elif part not in ("", "."):
            parts.append(part)
    if not parts:
        raise ValueError("Empty audio resource path")
    return "/".join(parts)


def read_member(archive, name, maximum):
    info = archive.getinfo(name)
    if info.file_size > maximum or info.flag_bits & 1:
        raise ValueError("Oversized or encrypted ZIP member: " + name)
    with archive.open(info) as stream:
        data = stream.read(maximum + 1)
    if len(data) > maximum:
        raise ValueError("ZIP member exceeded its read bound: " + name)
    return data


def read_xml(archive, name):
    data = read_member(archive, name, 2 * 1024 * 1024)
    # UTF-16 declarations remain visible after null removal. External entities are
    # never resolved; rejecting DTDs also excludes internal entity expansion.
    if b"<!DOCTYPE" in data.replace(b"\x00", b"").upper() or b"<!ENTITY" in data.replace(b"\x00", b"").upper():
        raise UncheckedAudio("DTD/entity XML is not checked by the audio evidence lane")
    root = ET.fromstring(data)
    if any("{http://www.w3.org/XML/1998/namespace}base" in node.attrib for node in root.iter()):
        raise UncheckedAudio("XML base addressing is not checked by the audio evidence lane")
    return root


def collect_clips(archive):
    names = archive.namelist()
    if len(names) > 10000 or len(names) != len(set(names)):
        raise ValueError("ZIP entries exceed the count limit or contain duplicate names")
    if "META-INF/encryption.xml" in names:
        raise UncheckedAudio("Encrypted/obfuscated publications require a separate audio-resource assessment")
    container = read_xml(archive, "META-INF/container.xml")
    if container.tag != CONTAINER + "container":
        raise ValueError("Invalid EPUB container root")
    roots = container.findall(CONTAINER + "rootfiles/" + CONTAINER + "rootfile")
    if len(roots) != 1:
        raise UncheckedAudio("Audio evidence requires one declared package rendition")
    package_path = local_path("", roots[0].get("full-path", ""))
    package = read_xml(archive, package_path)
    if package.tag != OPF + "package":
        raise ValueError("Invalid package root")
    resources = {}
    for item in package.findall(OPF + "manifest/" + OPF + "item"):
        # Remote non-audio resources do not enter this local-only lane.
        media = item.get("media-type", "")
        if media != "application/smil+xml" and not media.startswith("audio/"):
            continue
        path = local_path(package_path, item.get("href", ""))
        if path in resources:
            raise ValueError("Ambiguous manifest resource: " + path)
        resources[path] = media
    clips = []
    overlays = [path for path, media in resources.items() if media == "application/smil+xml"]
    if len(overlays) > 256:
        raise ValueError("Audio evidence supports at most 256 overlays")
    for path in overlays:
        smil = read_xml(archive, path)
        if smil.tag != SMIL + "smil":
            raise ValueError("Invalid SMIL root")
        for audio in smil.iter(SMIL + "audio"):
            target = local_path(path, audio.get("src", ""))
            if target not in resources or not resources[target].startswith("audio/"):
                raise ValueError("Audio target lacks an audio manifest declaration: " + target)
            end = audio.get("clipEnd")
            if end is None:
                raise UncheckedAudio("An audio clip without clipEnd requires separate implicit-timing qualification")
            begin = clock(audio.get("clipBegin", "0"))
            end = clock(end)
            if not Decimal(0) <= begin < end:
                raise ValueError("An audio clip has invalid begin/end ordering")
            clips.append({"index": len(clips) + 1, "boundsStatus": "not-checked", "overlay": path, "audio": target, "beginSeconds": str(begin), "endSeconds": str(end), "mediaType": resources[target]})
            if len(clips) > 10000:
                raise ValueError("Audio evidence supports at most 10000 clips")
    return clips


def measured_duration(report):
    streams = report.get("streams") if isinstance(report, dict) else None
    if not isinstance(streams, list) or len(streams) != 1 or not isinstance(streams[0], dict) or streams[0].get("codec_type") != "audio":
        raise UncheckedAudio("Exactly one audio stream is required for duration evidence")
    # Prefer the selected audio stream, not a possibly longer multiplexed container.
    value = streams[0].get("duration")
    if value in (None, "N/A"):
        container = report.get("format")
        value = container.get("duration") if isinstance(container, dict) else None
    try:
        duration = Decimal(str(value))
    except InvalidOperation as error:
        raise UncheckedAudio("No usable encoded-audio duration") from error
    if not duration.is_finite() or duration <= 0:
        raise UncheckedAudio("No finite positive encoded-audio duration")
    return duration


def inspect_audio(snapshot, folder, ffprobe, ffmpeg, timeout):
    evidence = {"status": "failed", "scope": "explicit-clip-bounds-against-probed-audio-duration",
                "nativePlayback": "not-checked", "decode": "not-checked", "clips": [], "resources": []}
    try:
        with zipfile.ZipFile(snapshot) as archive:
            clips = collect_clips(archive)
            evidence["clips"] = clips
            if not clips:
                evidence["status"] = "not-applicable"
                return evidence
            unique = list(dict.fromkeys(clip["audio"] for clip in clips))
            if len(unique) > 256:
                raise ValueError("Audio evidence supports at most 256 resources")
            total = 0
            with tempfile.TemporaryDirectory(prefix="audio-input-", dir=folder) as scratch:
                for index, resource in enumerate(unique, 1):
                    media = next(clip["mediaType"] for clip in clips if clip["audio"] == resource)
                    if media not in DEMUXERS:
                        raise UncheckedAudio("Unsupported audio media type: " + media)
                    data = read_member(archive, resource, 128 * 1024 * 1024)
                    total += len(data)
                    if total > 256 * 1024 * 1024:
                        raise ValueError("Expanded audio exceeds the 256 MiB publication limit")
                    local = Path(scratch) / (str(index) + ".media")
                    local.write_bytes(data)
                    row = {"path": resource, "sha256": hashlib.sha256(data).hexdigest(), "status": "failed", "decode": "not-checked"}
                    evidence["resources"].append(row)
                    del data
                    # Fixed demuxers exclude playlist autodetection. No network protocols;
                    # MOV external data references remain disabled explicitly.
                    inputs = ["-protocol_whitelist", "file,pipe", "-f", DEMUXERS[media]]
                    if DEMUXERS[media] == "mov":
                        inputs += ["-enable_drefs", "0", "-use_absolute_path", "0"]
                    inputs += ["-i", str(local)]
                    report_path = folder / f"audio-{index:04d}.ffprobe.json"
                    with report_path.open("wb") as output, (folder / f"audio-{index:04d}.ffprobe.log").open("wb") as log:
                        result = subprocess.run([str(ffprobe), "-v", "error"] + inputs +
                            ["-show_entries", "stream=codec_type,codec_name,duration,start_time:format=duration", "-of", "json"],
                            stdout=output, stderr=log, timeout=timeout, check=False)
                    row["probeExitCode"] = result.returncode
                    row["report"] = str(report_path.relative_to(folder.parent))
                    row["probeLog"] = str((folder / f"audio-{index:04d}.ffprobe.log").relative_to(folder.parent))
                    if result.returncode or report_path.stat().st_size > 1024 * 1024:
                        raise ValueError("Audio probe failed or exceeded its report bound")
                    duration = measured_duration(json.loads(report_path.read_text(encoding="utf-8")))
                    row["durationSeconds"] = str(duration)
                    for clip in clips:
                        if clip["audio"] == resource:
                            clip["boundsStatus"] = "failed" if Decimal(clip["endSeconds"]) > duration else "passed"
                    row["outOfBoundsClips"] = sum(clip["boundsStatus"] == "failed" for clip in clips if clip["audio"] == resource)
                    row["status"] = "failed" if row["outOfBoundsClips"] else "passed"
                    if ffmpeg:
                        with (folder / f"audio-{index:04d}.decode.log").open("wb") as log:
                            result = subprocess.run([str(ffmpeg), "-nostdin", "-v", "error", "-xerror"] + inputs +
                                ["-map", "0:a:0", "-f", "null", "-"], stdout=subprocess.DEVNULL, stderr=log,
                                timeout=timeout, check=False)
                        row["decodeLog"] = str((folder / f"audio-{index:04d}.decode.log").relative_to(folder.parent))
                        row["decodeExitCode"] = result.returncode
                        row["decode"] = "passed" if result.returncode == 0 else "failed"
                        if result.returncode:
                            row["status"] = "failed"
                    local.unlink()
            evidence["status"] = "passed" if all(row["status"] == "passed" for row in evidence["resources"]) else "failed"
            if ffmpeg:
                evidence["decode"] = "passed" if all(row["decode"] == "passed" for row in evidence["resources"]) else "failed"
    except UncheckedAudio as error:
        evidence["status"] = "failed" if any(row.get("outOfBoundsClips", 0) or row.get("decode") == "failed" for row in evidence["resources"]) else "not-checked"
        evidence["reason"] = str(error)
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile, NotImplementedError, subprocess.SubprocessError) as error:
        evidence["reason"] = str(error)
    return evidence
