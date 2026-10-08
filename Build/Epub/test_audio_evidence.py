import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from audio_evidence import collect_clips, clock, inspect_audio, local_path, measured_duration, UncheckedAudio


class AudioEvidenceTests(unittest.TestCase):
    def publication(self, clip_end="2.5s", clip_begin="0s", extras=None, src="../audio/read.mp3"):
        stream = io.BytesIO()
        entries = {
            "META-INF/container.xml": '<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="EPUB/package.opf"/></rootfiles></container>',
            "EPUB/package.opf": '<package xmlns="http://www.idpf.org/2007/opf"><manifest><item id="mo" href="overlays/read.smil" media-type="application/smil+xml"/><item id="a" href="audio/read.mp3" media-type="audio/mpeg"/></manifest></package>',
            "EPUB/overlays/read.smil": '<smil xmlns="http://www.w3.org/ns/SMIL"><body><seq><seq><par><audio src="' + src + '" clipBegin="' + clip_begin + '"' + ('' if clip_end is None else ' clipEnd="' + clip_end + '"') + '/></par></seq></seq></body></smil>',
            "EPUB/audio/read.mp3": b"test audio bytes"
        }
        entries.update(extras or {})
        with zipfile.ZipFile(stream, "w") as archive:
            for name, content in entries.items():
                archive.writestr(name, content)
        return stream.getvalue()

    def test_nested_clips_and_exact_clock_values(self):
        with zipfile.ZipFile(io.BytesIO(self.publication("00:00:02.5000001", "25ms"))) as archive:
            clips = collect_clips(archive)
        self.assertEqual("EPUB/audio/read.mp3", clips[0]["audio"])
        self.assertEqual("2.5000001", clips[0]["endSeconds"])
        self.assertEqual("0.025", clips[0]["beginSeconds"])
        self.assertEqual(clock("1min"), clock("60s"))
        for value in ("NaN", "-1s", "1e2", "00:60:00", "00:00:60"):
            with self.assertRaises(ValueError):
                clock(value)

    def test_missing_end_and_remote_resources_are_unchecked(self):
        for data in (self.publication(None), self.publication(src="https://example.test/a.mp3")):
            with zipfile.ZipFile(io.BytesIO(data)) as archive, self.assertRaises(UncheckedAudio):
                collect_clips(archive)

    def test_ambiguous_and_escaping_paths_are_rejected(self):
        for reference in ("../../../outside.mp3", "audio%2fread.mp3", "a\\b.mp3", "audio%00.mp3"):
            with self.assertRaises(ValueError):
                local_path("EPUB/read.smil", reference)
        self.assertEqual("EPUB/audio/a b.mp3", local_path("EPUB/mo/read.smil", "../audio/a%20b.mp3"))

    def test_stream_duration_wins_over_longer_container(self):
        self.assertEqual(clock("2s"), measured_duration({"streams": [{"codec_type": "audio", "duration": "2"}], "format": {"duration": "100"}}))
        for report in ({}, {"streams": [None]}, {"streams": [{"codec_type": "audio", "duration": "NaN"}]},
                       {"streams": [{"codec_type": "audio"}], "format": None}):
            with self.assertRaises(UncheckedAudio):
                measured_duration(report)

    def test_overrun_is_failed_and_probe_evidence_is_retained(self):
        def fake_probe(command, stdout, **kwargs):
            self.assertIn("-protocol_whitelist", command)
            self.assertIn("mp3", command)
            stdout.write(json.dumps({"streams": [{"codec_type": "audio", "duration": "2"}]}).encode())
            return type("Result", (), {"returncode": 0})()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "book.epub"; source.write_bytes(self.publication())
            with patch("audio_evidence.subprocess.run", side_effect=fake_probe):
                result = inspect_audio(source, root, Path("ffprobe"), None, 10)
            self.assertEqual("failed", result["status"])
            self.assertEqual(1, result["resources"][0]["outOfBoundsClips"])
            self.assertEqual("failed", result["clips"][0]["boundsStatus"])
            self.assertEqual("not-checked", result["decode"])
            self.assertTrue((root.parent / result["resources"][0]["report"]).is_file())
            self.assertFalse(list(root.glob("audio-input-*")))

    def test_decode_failure_cannot_pass_duration_success(self):
        def fake_process(command, stdout, **kwargs):
            if "-show_entries" in command:
                stdout.write(json.dumps({"streams": [{"codec_type": "audio", "duration": "3"}]}).encode())
                code = 0
            else:
                self.assertIn("-xerror", command)
                code = 1
            return type("Result", (), {"returncode": code})()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory); source = root / "book.epub"; source.write_bytes(self.publication())
            with patch("audio_evidence.subprocess.run", side_effect=fake_process):
                result = inspect_audio(source, root, Path("ffprobe"), Path("ffmpeg"), 10)
            self.assertEqual("failed", result["status"])
            self.assertEqual("failed", result["decode"])
            self.assertEqual("not-checked", result["nativePlayback"])

    def test_xml_addressing_and_entities_are_not_silently_ignored(self):
        for xml in ('<smil xmlns="http://www.w3.org/ns/SMIL" xml:base="other/"/>',
                    '<!DOCTYPE smil [<!ENTITY a "text">]><smil xmlns="http://www.w3.org/ns/SMIL"/>'):
            with zipfile.ZipFile(io.BytesIO(self.publication(extras={"EPUB/overlays/read.smil": xml}))) as archive, self.assertRaises(UncheckedAudio):
                collect_clips(archive)


if __name__ == "__main__":
    unittest.main()
