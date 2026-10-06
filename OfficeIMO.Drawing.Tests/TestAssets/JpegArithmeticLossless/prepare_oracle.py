"""Configure an isolated thorfdbg/libjpeg checkout for fixture generation.

Use commit c719010a26ce0c666e98b2acf924ad5fc24b4f5d. These edits expose existing
scan settings in the test driver and correct the initial predictor per T.81
H.1.2.1. No entropy-coding implementation is changed or copied into OfficeIMO.
Build the executable with ./configure && make -j4 final after running this script.
"""
import pathlib
import subprocess
import sys

root = pathlib.Path(sys.argv[1]).resolve()
assert subprocess.check_output(['git', '-C', str(root), 'rev-parse', 'HEAD'], text=True).strip() == 'c719010a26ce0c666e98b2acf924ad5fc24b4f5d'
edits = {
 'cmd/encodec.cpp': [('            JPG_ValueTag(JPGTAG_IMAGE_FRAMETYPE,frametype),',
 '''            JPG_ValueTag(JPGTAG_IMAGE_FRAMETYPE,frametype),
            JPG_ValueTag(JPGTAG_SCAN_POINTTRANSFORM,getenv("OFFICEIMO_TEST_POINT") ? atoi(getenv("OFFICEIMO_TEST_POINT")) : 0),
            JPG_ValueTag(JPGTAG_SCAN_SPECTRUM_START,getenv("OFFICEIMO_TEST_PREDICTOR") ? atoi(getenv("OFFICEIMO_TEST_PREDICTOR")) : 4),''')],
 'marker/scan.cpp': [('m_ucScanStart = 4; // predictor to use. This is the default.',
 'm_ucScanStart = tags->GetTagData(JPGTAG_SCAN_SPECTRUM_START,4); // test harness selects the standard predictor.')],
 'codestream/predictivescan.cpp': [('FractionalColorBitsOf() + m_ucLowBit,(1L << m_pFrame->PrecisionOf()) >> 1);',
 'FractionalColorBitsOf() + m_ucLowBit,(1L << (m_pFrame->PrecisionOf() - m_ucLowBit)) >> 1);')]
}
for name, replacements in edits.items():
 path = root / name
 text = path.read_text()
 for old, new in replacements:
  if new in text: continue
  assert text.count(old) == 1, name
  text = text.replace(old, new)
 path.write_text(text)
