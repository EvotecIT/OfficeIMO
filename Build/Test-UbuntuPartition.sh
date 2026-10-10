#!/usr/bin/env bash
# Runs one repository test partition with the workflow-provided matrix environment.
# The caller uses bash -e to preserve the existing CI failure behavior.
# Keep the heavy Excel image visual gate on its dedicated script path; run test projects
# explicitly so Ubuntu does not spend the job budget walking every solution project.
test_filter='Category!=ExcelImageVisualGate'
run_shared_tests() {
  local name="$1"
  local collect_coverage="${2:-false}"
  local coverage_args=()
  # Keep PR test partitions focused on contract proof. Coverage collection is
  # intentionally disabled for the matrix because it turns data-heavy Markdown
  # and PDF slices into long-running hosted-runner jobs.
  if [ "${COLLECT_COVERAGE}" = "true" ] && [ "${collect_coverage}" = "true" ]; then
    coverage_args=(--collect "XPlat Code Coverage")
  fi
  echo ""
  echo "== OfficeIMO.Shared.Tests: ${name} =="
  dotnet test OfficeIMO.Shared.Tests/OfficeIMO.Shared.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx "${coverage_args[@]}"
}

run_excel_tests() {
  local name="$1"
  local filter="$2"
  echo ""
  echo "== OfficeIMO.Excel.Tests: ${name} =="
  dotnet test OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "(${filter})&${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_word_tests() {
  local name="$1"
  local filter="$2"
  echo ""
  echo "== OfficeIMO.Word.Tests: ${name} =="
  dotnet test OfficeIMO.Word.Tests/OfficeIMO.Word.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "(${filter})&${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_markdown_tests() {
  local name="$1"
  local filter="$2"
  echo ""
  echo "== OfficeIMO.Markdown.Tests: ${name} =="
  dotnet test OfficeIMO.Markdown.Tests/OfficeIMO.Markdown.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "(${filter})&${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_powerpoint_tests() {
  echo ""
  echo "== OfficeIMO.PowerPoint.Tests =="
  dotnet test OfficeIMO.PowerPoint.Tests/OfficeIMO.PowerPoint.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_visio_tests() {
  echo ""
  echo "== OfficeIMO.Visio.Tests =="
  dotnet test OfficeIMO.Visio.Tests/OfficeIMO.Visio.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_access_tests() {
  dotnet test OfficeIMO.Access.Tests/OfficeIMO.Access.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_rtf_tests() {
  echo ""
  echo "== OfficeIMO.Rtf.Tests =="
  dotnet test OfficeIMO.Rtf.Tests/OfficeIMO.Rtf.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_html_tests() {
  echo ""
  echo "== OfficeIMO.Html.Tests =="
  dotnet test OfficeIMO.Html.Tests/OfficeIMO.Html.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_reader_tests() {
  echo ""
  echo "== OfficeIMO.Reader.Tests =="
  dotnet test OfficeIMO.Reader.Tests/OfficeIMO.Reader.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_email_tests() {
  echo ""
  echo "== OfficeIMO.Email.Tests =="
  dotnet test OfficeIMO.Email.Tests/OfficeIMO.Email.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "Category!=Performance" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_pdf_tests() {
  local name="$1"
  local filter="$2"
  echo ""
  echo "== OfficeIMO.Pdf.Tests: ${name} =="
  dotnet test OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "(${filter})&${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_drawing_tests() {
  echo ""
  echo "== OfficeIMO.Drawing.Tests =="
  dotnet test OfficeIMO.Drawing.Tests/OfficeIMO.Drawing.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_googleworkspace_tests() {
  echo ""
  echo "== OfficeIMO.GoogleWorkspace.Tests =="
  dotnet test OfficeIMO.GoogleWorkspace.Tests/OfficeIMO.GoogleWorkspace.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_markup_tests() {
  echo ""
  echo "== OfficeIMO.Markup.Tests =="
  dotnet test OfficeIMO.Markup.Tests/OfficeIMO.Markup.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

run_opendocument_tests() {
  echo ""
  echo "== OfficeIMO.OpenDocument.Tests =="
  dotnet test OfficeIMO.OpenDocument.Tests/OfficeIMO.OpenDocument.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
}

audit_retired_aggregate_tests() {
  echo ""
  echo "== retired aggregate test project audit =="
  if [ -f OfficeIMO.Tests/OfficeIMO.Tests.csproj ]; then
    echo "OfficeIMO.Tests/OfficeIMO.Tests.csproj must stay retired. Put tests in a domain project or OfficeIMO.Shared.Tests."
    exit 1
  fi
  if find OfficeIMO.Tests -maxdepth 1 -name '*.cs' -print -quit 2>/dev/null | grep -q .; then
    echo "OfficeIMO.Tests still contains root test sources. Move them to a domain test project or OfficeIMO.Shared.Tests."
    exit 1
  fi
}

audit_markdown_test_partitions() {
  local list_file
  list_file="$(mktemp)"
  echo ""
  echo "== OfficeIMO.Markdown.Tests: partition coverage audit =="
  dotnet test OfficeIMO.Markdown.Tests/OfficeIMO.Markdown.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity quiet --framework "${TEST_FRAMEWORK}" --list-tests --filter "${test_filter}" > "${list_file}"
  python3 - "${list_file}" <<'PY'
import sys

with open(sys.argv[1], encoding="utf-8", errors="replace") as handle:
    discovered = [line.strip() for line in handle if line.startswith("    ") and line.strip()]

markdown_partitions = (
    "OfficeIMO.Tests.Markdown.",
    "OfficeIMO.Tests.MarkdownSuite",
    "OfficeIMO.Tests.MarkdownHtmlToMarkdownCompatibilityTests",
    "OfficeIMO.Tests.MarkdownHtmlToMarkdownTests",
    "OfficeIMO.Tests.MarkdownHtmlApiContracts",
    "OfficeIMO.Tests.Markdown_",
    "OfficeIMO.Tests.MarkdownTranscript",
    "OfficeIMO.Tests.MarkdownReaderBomTests",
    "OfficeIMO.Tests.MarkdownBlankLinesTests",
    "OfficeIMO.Tests.MarkdownDefinitionLinesTests",
    "OfficeIMO.Tests.MarkdownOrderedListsTests",
    "OfficeIMO.Tests.MarkdownStreamingPreviewTests",
    "OfficeIMO.Tests.Markdown_Renderer",
    "OfficeIMO.Tests.Markdown_IntelligenceX",
    "OfficeIMO.Tests.MarkdownAllSeverityBatch11SecurityTests",
    "OfficeIMO.Tests.MarkdownAllSeverityBatch12SecurityTests",
    "OfficeIMO.Tests.MarkdownAllSeverityBatch15SecurityTests",
    "OfficeIMO.Tests.MarkdownAllSeverityBatch21Tests",
)

unassigned = [name for name in discovered if not any(partition in name for partition in markdown_partitions)]
if unassigned:
    print("OfficeIMO.Markdown.Tests partition audit found tests that are not assigned to a named CI slice:")
    for name in unassigned:
        print(f"  {name}")
    sys.exit(1)

print(f"OfficeIMO.Markdown.Tests partition audit covered {len(discovered)} discovered tests.")
PY
  rm -f "${list_file}"
}

if [ "${TEST_FRAMEWORK}" = "net8.0" ] && [ "${TEST_PARTITION}" = "markdown-large" ]; then
  audit_markdown_test_partitions
fi

case "${TEST_PARTITION}" in
  markdown-large)
    run_word_tests "word-markdown-roundtrip" 'FullyQualifiedName~OfficeIMO.Tests.MarkdownMoreCasesTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownRoundTripTests'
    run_markdown_tests "markdown-root" 'FullyQualifiedName~OfficeIMO.Tests.Markdown.'
    ;;
  markdown-suite)
    run_markdown_tests "markdown-suite" 'FullyQualifiedName~OfficeIMO.Tests.MarkdownSuite'
    run_markdown_tests "markdown-named" 'FullyQualifiedName~OfficeIMO.Tests.MarkdownHtmlToMarkdownCompatibilityTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownHtmlToMarkdownTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownHtmlApiContracts|FullyQualifiedName~OfficeIMO.Tests.Markdown_|FullyQualifiedName~OfficeIMO.Tests.MarkdownTranscript|FullyQualifiedName~OfficeIMO.Tests.MarkdownReaderBomTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownBlankLinesTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownDefinitionLinesTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownOrderedListsTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownStreamingPreviewTests|FullyQualifiedName~OfficeIMO.Tests.Markdown_Renderer|FullyQualifiedName~OfficeIMO.Tests.Markdown_IntelligenceX|FullyQualifiedName~OfficeIMO.Tests.MarkdownAllSeverityBatch11SecurityTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownAllSeverityBatch12SecurityTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownAllSeverityBatch15SecurityTests|FullyQualifiedName~OfficeIMO.Tests.MarkdownAllSeverityBatch21Tests'
    ;;
  pdf-visual-inspector)
    export OFFICEIMO_REQUIRE_PDF_RASTERIZER=1
    # The committed pixel baselines are authoring-platform evidence.
    # Ubuntu still renders every scenario with the pinned Poppler build,
    # but host font substitution makes cross-platform pixel identity invalid.
    export OFFICEIMO_PDF_RASTER_SMOKE_ONLY=1
    run_pdf_tests "pdf-visual-inspector" 'FullyQualifiedName~OfficeIMO.Tests.Pdf.PdfDocumentVisualQualityTests|FullyQualifiedName~OfficeIMO.Tests.Pdf.PdfInspectorTests|FullyQualifiedName~OfficeIMO.Tests.Pdf.PdfReaderAndFooterRegressionTests|FullyQualifiedName~OfficeIMO.Tests.Pdf.MarkdownSaveAsPdfVisualTests|FullyQualifiedName~OfficeIMO.Tests.Pdf.PdfDocumentRasterVisualBaselineTests'
    ;;
  pdf-core)
    # Every PDF test outside the visual partition belongs to core, regardless of namespace.
    run_pdf_tests "pdf-core" 'FullyQualifiedName!~OfficeIMO.Tests.Pdf.PdfDocumentVisualQualityTests&FullyQualifiedName!~OfficeIMO.Tests.Pdf.PdfInspectorTests&FullyQualifiedName!~OfficeIMO.Tests.Pdf.PdfReaderAndFooterRegressionTests&FullyQualifiedName!~OfficeIMO.Tests.Pdf.MarkdownSaveAsPdfVisualTests&FullyQualifiedName!~OfficeIMO.Tests.Pdf.PdfDocumentRasterVisualBaselineTests'
    ;;
  excel-image-charts)
    run_excel_tests "excel-image" 'FullyQualifiedName~OfficeIMO.Tests.ExcelImageExport'
    run_excel_tests "excel-charts" 'FullyQualifiedName~OfficeIMO.Tests.Excel.Test_ExcelChart|FullyQualifiedName~OfficeIMO.Tests.Excel.SharedChartRenderer'
    run_excel_tests "excel-closedxml-gap" 'FullyQualifiedName~OfficeIMO.Tests.Excel.Test_ClosedXmlGap'
    ;;
  excel-legacy-reader)
    dotnet test OfficeIMO.LegacyImport.Tests/OfficeIMO.LegacyImport.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --no-restore --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    run_excel_tests "excel-legacy-xls" 'FullyQualifiedName~OfficeIMO.Tests.Excel.LegacyXls'
    run_excel_tests "excel-performance" 'FullyQualifiedName~OfficeIMO.Tests.Excel.PerformanceReview'
    run_excel_tests "excel-reader" 'FullyQualifiedName~OfficeIMO.Tests.Excel.Reader|FullyQualifiedName~OfficeIMO.Tests.Excel.ReadRows|FullyQualifiedName~OfficeIMO.Tests.Excel.ExcelHttp'
    run_excel_tests "excel-save-pdf" 'FullyQualifiedName~OfficeIMO.Tests.Excel.SaveAsPdf|FullyQualifiedName~OfficeIMO.Tests.Excel.ToPdfDocument|FullyQualifiedName~OfficeIMO.Tests.Excel.PdfTables|FullyQualifiedName~OfficeIMO.Tests.Excel.ReportWorkflow'
    ;;
  excel-core-named)
    run_excel_tests "excel-named-typed" 'FullyQualifiedName~OfficeIMO.Tests.ExcelCellAndRangeTests|FullyQualifiedName~OfficeIMO.Tests.ExcelDelimitedImportTests|FullyQualifiedName~OfficeIMO.Tests.ExcelFluent|FullyQualifiedName~OfficeIMO.Tests.ExcelInfoBuilderTests|FullyQualifiedName~OfficeIMO.Tests.ExcelNamed|FullyQualifiedName~OfficeIMO.Tests.ExcelPowerShell|FullyQualifiedName~OfficeIMO.Tests.ExcelRange|FullyQualifiedName~OfficeIMO.Tests.ExcelRows|FullyQualifiedName~OfficeIMO.Tests.ExcelSheet|FullyQualifiedName~OfficeIMO.Tests.ExcelCoerce|FullyQualifiedName~OfficeIMO.Tests.ExcelAllSeverityBatch16SecurityTests'
    run_excel_tests "excel-core" '(FullyQualifiedName~OfficeIMO.Tests.Excel.&FullyQualifiedName!~OfficeIMO.Tests.Excel.Test_ExcelChart&FullyQualifiedName!~OfficeIMO.Tests.Excel.SharedChartRenderer&FullyQualifiedName!~OfficeIMO.Tests.Excel.Test_ClosedXmlGap&FullyQualifiedName!~OfficeIMO.Tests.Excel.LegacyXls&FullyQualifiedName!~OfficeIMO.Tests.Excel.PerformanceReview&FullyQualifiedName!~OfficeIMO.Tests.Excel.SaveAsPdf&FullyQualifiedName!~OfficeIMO.Tests.Excel.ToPdfDocument&FullyQualifiedName!~OfficeIMO.Tests.Excel.PdfTables&FullyQualifiedName!~OfficeIMO.Tests.Excel.ReportWorkflow&FullyQualifiedName!~OfficeIMO.Tests.Excel.Reader&FullyQualifiedName!~OfficeIMO.Tests.Excel.ReadRows&FullyQualifiedName!~OfficeIMO.Tests.Excel.ExcelHttp)'
    ;;
  word-rtf-html)
    run_word_tests "word" 'FullyQualifiedName~OfficeIMO.Tests.Word|FullyQualifiedName~OfficeIMO.Tests.ImageEmbedderTests|FullyQualifiedName~OfficeIMO.Tests.ParagraphParentMutationTests'
    run_rtf_tests
    run_html_tests
    ;;
  email)
    run_email_tests
    ;;
  other-projects)
    audit_retired_aggregate_tests
    dotnet test OfficeIMO.Adf.Tests/OfficeIMO.Adf.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.DocBook.Tests/OfficeIMO.DocBook.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.Confluence.Tests/OfficeIMO.Confluence.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    run_drawing_tests
    dotnet test OfficeIMO.Drawing.CodeGlyphX.Tests/OfficeIMO.Drawing.CodeGlyphX.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    run_googleworkspace_tests
    dotnet test OfficeIMO.IWork.Tests/OfficeIMO.IWork.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    run_markup_tests
    dotnet test OfficeIMO.Xps.Tests/OfficeIMO.Xps.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.DjVu.Tests/OfficeIMO.DjVu.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    run_opendocument_tests
    dotnet test OfficeIMO.OpenDocument.Converters.Tests/OfficeIMO.OpenDocument.Converters.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --filter "${test_filter}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    run_powerpoint_tests
    run_visio_tests
    run_access_tests
    run_reader_tests
    dotnet test OfficeIMO.OneNote.Tests/OfficeIMO.OneNote.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    run_shared_tests "shared-contracts"
    dotnet test OfficeIMO.ChartForgeX.Tests/OfficeIMO.ChartForgeX.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.ChartForgeX.Markdown.Tests/OfficeIMO.ChartForgeX.Markdown.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" -m:1 --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.Tool.Tests/OfficeIMO.Tool.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.Bibliography.Tests/OfficeIMO.Bibliography.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.CSV.Tests/OfficeIMO.CSV.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.VerifyTests/OfficeIMO.VerifyTests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.MarkdownRenderer.Wpf.Tests/OfficeIMO.MarkdownRenderer.Wpf.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.AsciiDoc.Tests/OfficeIMO.AsciiDoc.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.AsciiDoc.Markdown.Tests/OfficeIMO.AsciiDoc.Markdown.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.Latex.Tests/OfficeIMO.Latex.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    dotnet test OfficeIMO.Latex.Markdown.Tests/OfficeIMO.Latex.Markdown.Tests.csproj --configuration "${BUILD_CONFIGURATION}" --no-build --verbosity "${TEST_VERBOSITY}" --framework "${TEST_FRAMEWORK}" --logger "console;verbosity=${TEST_VERBOSITY}" --logger trx
    ;;
  *)
    echo "Unknown partition: ${TEST_PARTITION}"
    exit 1
    ;;
esac
