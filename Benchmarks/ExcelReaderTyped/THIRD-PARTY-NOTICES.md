The typed/raw fixture values, writer operations, record shape, accumulators, and
BIFF8/OLE and string-heavy fixture generators follow ExcelReader's benchmarks at commit
ca5b50f99e8ef57ab476f0a2bc8043558d58b28d.
The implementation here adds shape and output qualification. Optional CSV and
Arrow cases also adapt ArrowConversionBenchmark.cs and Shared/CsvGenerator.cs
from that same pinned source revision.
The model and lifecycle cases also follow Parse/ParseBenchmark.cs,
RealData/RealDataTypedParseBenchmark.cs, Read/XlsxSharedStringHotPathBenchmark.cs,
and Write/WritePathBenchmark.cs from that revision.
The encrypted diagnostics follow Crypto/EncryptedWorkbookBenchmark.cs. Small
encrypted fixture bytes and their paired plaintext oracle remain in the
separately attributed OfficeIMO.TestAssets/Documents/ExcelEncryptionCorpus.
The real CSV and record-writing cases also follow
RealData/RealDataReadBenchmark.cs and Write/RecordWriteBenchmark.cs. Real CSV
data is hash-pinned and downloaded for qualification; this project does not
redistribute that fixture.

Source: https://github.com/GabrielMarquezMatte/ExcelReader

MIT License

Copyright (c) 2026 Gabriel Matte

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
