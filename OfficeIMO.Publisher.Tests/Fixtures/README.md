# Publisher native fixtures

These independent-producer publications come from Apache POI's
[`test-data/publisher`](https://github.com/apache/poi/tree/66109187d12b72696ea02101cec765d77d9bbdc9/test-data/publisher)
corpus. Source revision and file hashes are recorded in `provenance.json`.
`LICENSE-APACHE-POI.txt` and `NOTICE-APACHE-POI.txt` retain the upstream terms.

Simple, Sample and Sample_2010 exercise positioned text, styles and tables.
SampleBrochure and SampleNewsletter exercise grouped artwork, GIF/JPEG/WMF
resources, page order and linked stories. Sample2000 and Sample98 protect the
explicit earlier-generation rejection contract.

The fixtures are test inputs, not native output authored by OfficeIMO. Expected
simple/table geometry and style values are checked against libmspub 0.1.5's
`pub2raw` output. Native Publisher PDF renderings are not supplied by this corpus.
