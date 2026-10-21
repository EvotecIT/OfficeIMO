# Font parser resource fixtures

These original test fonts were generated with fontTools 4.67.0 by CodeGlyphX's
`Build/Generate-PrintFontFixtures.py` (source revision
`8548c1a95935a1c758bd4ea17474de974dbc831f`). Copyright EvotecIT; Apache License 2.0
(included in `LICENSE-Apache-2.0.txt`). They contain simple test outlines and
are embedded only in the test assembly.

- `TestSharedCff2.otf.gz`: 256 font dictionaries reference one 16,384-entry
  subroutine index. Shared metadata must remain shared when decoded.
- `TestCumulativeCff2.otf.gz`: distinct subroutine indices exceed the cumulative
  metadata allowance and must be rejected.
- `TestSubstitutionCoverage.ttf.gz`: repeated nonmatching coverage ranges make
  the range-search work exceed the shaping allowance.

Each fixture expands to less than 3 MiB. Tests exercise deterministic parser
limits, without allocation thresholds, elapsed-time assertions, or large stress
inputs. fontTools is a fixture-authoring tool, not a build or runtime dependency.
