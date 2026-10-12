/*
 * The engine's document journey from a native host: open, read facts, render, rotate, undo, redo, organize pages,
 * extract, search and mark up, save a copy and reopen it, then make a searchable PDF from a scanned page image. A .NET process cannot host the library on macOS or Linux, because its runtime and the library's handle
 * the same thread-suspension signal, so the journey runs from C like the native apps do.
 *
 * Usage: officeimo_journey <library> <pdf> <work-directory>
 * The PDF is opened in place and must be a copy; the rendered page and the saved copy are written to the work
 * directory. Exit code 0 means every step passed.
 *
 * The library is loaded at run time, and every entry point takes its type from officeimo_studio.h, so the header and
 * the exports are checked against each other.
 */
#include "officeimo_studio.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>

#ifdef _WIN32
#include <windows.h>
typedef HMODULE library_handle;
static library_handle load_library(const char *path) { return LoadLibraryA(path); }
static void *load_symbol(library_handle library, const char *name) { return (void *)GetProcAddress(library, name); }
#define PATH_SEPARATOR "\\"
#else
#include <dlfcn.h>
typedef void *library_handle;
static library_handle load_library(const char *path) { return dlopen(path, RTLD_NOW | RTLD_LOCAL); }
static void *load_symbol(library_handle library, const char *name) { return dlsym(library, name); }
#define PATH_SEPARATOR "/"
#endif

#if defined(_MSC_VER) && !defined(__clang__)
#define TYPE_OF(expression) typeof(expression)
#else
#define TYPE_OF(expression) __typeof__(expression)
#endif

#define ENGINE_FUNCTIONS(X) \
    X(officeimo_abi_version) X(officeimo_library_version) X(officeimo_free) X(officeimo_initialize) \
    X(officeimo_cancellation_create) X(officeimo_cancellation_cancel) X(officeimo_cancellation_free) \
    X(officeimo_document_open) X(officeimo_document_close) X(officeimo_document_info_json) \
    X(officeimo_document_render_page) X(officeimo_document_render_region) X(officeimo_document_rotate_pages) X(officeimo_document_undo) \
    X(officeimo_document_redo) X(officeimo_document_save) X(officeimo_document_delete_pages) \
    X(officeimo_document_duplicate_pages) X(officeimo_document_reorder_pages) X(officeimo_document_insert_blank_page) \
    X(officeimo_document_extract_pages) X(officeimo_document_search_json) X(officeimo_document_add_annotation_json) \
    X(officeimo_document_annotations_json) X(officeimo_document_update_annotation_json) X(officeimo_document_remove_annotation) \
    X(officeimo_document_page_text_json) X(officeimo_create_pdf_from_scan_json) X(officeimo_document_import_pages_json) \
    X(officeimo_document_make_searchable_json)

#define DECLARE_FUNCTION(name) static TYPE_OF(name) *engine_##name;
ENGINE_FUNCTIONS(DECLARE_FUNCTION)

static int failures;

static double milliseconds(void) {
    struct timespec now;
    timespec_get(&now, TIME_UTC);
    return now.tv_sec * 1000.0 + now.tv_nsec / 1000000.0;
}

static void check(const char *step, int32_t status, char *error, int32_t expected) {
    printf("%-28s status %d%s%s\n", step, status, error ? " - " : "", error ? error : "");
    if (error) engine_officeimo_free(error);
    if (status != expected) {
        printf("FAILED: %s returned %d, expected %d\n", step, status, expected);
        failures++;
    }
}

/* Runs a call, then checks its status and message. The call completes before error is read, and error is reset. */
#define CHECK(step, call, expected) \
    do { int32_t status_ = (call); check(step, status_, error, expected); error = NULL; } while (0)

static void expect(int condition, const char *description, const char *json) {
    if (!condition) {
        printf("FAILED: %s: %s\n", description, json ? json : "(no facts)");
        failures++;
    }
}

/* Reads the document facts. The caller releases the JSON with officeimo_free. */
static char *facts(officeimo_document document) {
    char *json = NULL, *error = NULL;
    CHECK("info", engine_officeimo_document_info_json(document, &json, &error), OFFICEIMO_OK);
    return json;
}

/* Finds "name": followed by value in the facts. */
static int has(const char *json, const char *name, const char *value) {
    char pattern[128];
    snprintf(pattern, sizeof pattern, "\"%s\":%s", name, value);
    return json && strstr(json, pattern) != NULL;
}

/* A stand-in for the platform's text recognizer: it checks it was given a PNG and reports one word on each page. */
static int32_t recognize_page(void *context, int32_t page, int32_t page_count, const uint8_t *png, size_t length, char **words_json) {
    static const char words[] = "[{\"text\":\"Journeyword\",\"x\":0.1,\"y\":0.1,\"width\":0.4,\"height\":0.05,\"confidence\":0.9}]";
    if (!png || length < 8 || memcmp(png + 1, "PNG", 3) != 0 || page < 1 || page > page_count) return OFFICEIMO_FAILED;
    *words_json = malloc(sizeof words);
    if (!*words_json) return OFFICEIMO_FAILED;
    memcpy(*words_json, words, sizeof words);
    (*(int *)context)++;
    return OFFICEIMO_OK;
}

/* Writes a path as a JSON string, escaping the backslashes of Windows paths and quotes. */
static void json_string(char *buffer, size_t size, const char *value) {
    size_t used = 0;
    if (size < 3) return;
    buffer[used++] = '"';
    for (; *value && used + 3 < size; value++) {
        if (*value == '\\' || *value == '"') buffer[used++] = '\\';
        buffer[used++] = *value;
    }
    buffer[used++] = '"';
    buffer[used] = '\0';
}

static void path_in(char *buffer, size_t size, const char *directory, const char *name) {
    snprintf(buffer, size, "%s%s%s", directory, PATH_SEPARATOR, name);
}

int main(int argc, char **argv) {
    if (argc != 4) {
        fprintf(stderr, "Usage: %s <library> <pdf> <work-directory>\n", argv[0]);
        return 2;
    }
    const char *library_path = argv[1], *source = argv[2], *work = argv[3];

    library_handle library = load_library(library_path);
    if (!library) {
        fprintf(stderr, "Could not load %s\n", library_path);
        return 1;
    }
#define LOAD_FUNCTION(name) \
    engine_##name = (TYPE_OF(name) *)load_symbol(library, #name); \
    if (!engine_##name) { fprintf(stderr, "Missing export %s\n", #name); return 1; }
    ENGINE_FUNCTIONS(LOAD_FUNCTION)

    char *version = engine_officeimo_library_version();
    printf("ABI version %d, engine %s\n", engine_officeimo_abi_version(), version ? version : "?");
    engine_officeimo_free(version);
    expect(engine_officeimo_abi_version() == OFFICEIMO_ABI_VERSION, "the library ABI differs from the header", NULL);

    char data[4096], page_png[4096], copy[4096];
    path_in(data, sizeof data, work, "data");
    path_in(page_png, sizeof page_png, work, "page-1.png");
    path_in(copy, sizeof copy, work, "rotated.pdf");

    char *error = NULL;
    officeimo_document document = 0;
    CHECK("open before initialize", engine_officeimo_document_open(source, NULL, 0, &document, &error),
          OFFICEIMO_NOT_INITIALIZED);
    CHECK("initialize", engine_officeimo_initialize(data, &error), OFFICEIMO_OK);

    double started = milliseconds();
    CHECK("open", engine_officeimo_document_open(source, NULL, 0, &document, &error), OFFICEIMO_OK);
    printf("  opened in %.0f ms\n", milliseconds() - started);
    if (!document) {
        printf("FAILED: no document handle\n");
        return 1;
    }
    char *json = facts(document);
    printf("  %s\n", json ? json : "");
    expect(has(json, "rotation", "0"), "page 1 is not upright", json);
    engine_officeimo_free(json);

    uint8_t *png = NULL;
    size_t length = 0;
    int32_t width = 0, height = 0;
    started = milliseconds();
    CHECK("render page 1", engine_officeimo_document_render_page(document, 1, 1.5, 0, &png, &length, &width, &height,
          &error), OFFICEIMO_OK);
    printf("  %d x %d px, %zu bytes in %.0f ms\n", width, height, length, milliseconds() - started);
    expect(png && length > 8 && memcmp(png, "\x89PNG", 4) == 0, "the page is not a PNG", NULL);
    if (png) {
        FILE *file = fopen(page_png, "wb");
        if (file) {
            fwrite(png, 1, length, file);
            fclose(file);
        }
        engine_officeimo_free(png);
    }
    /* The same page as display pixels: the whole page image, then its lower-right quarter as a tile. */
    uint8_t *pixels = NULL;
    size_t pixel_length = 0;
    started = milliseconds();
    CHECK("render region", engine_officeimo_document_render_region(document, 1, 1.5, 0, 0, width, height, 0, &pixels,
          &pixel_length, &error), OFFICEIMO_OK);
    printf("  %d x %d px in %.0f ms\n", width, height, milliseconds() - started);
    expect(pixels && pixel_length == (size_t)width * height * 4 && pixels[3] == 255, "the region is not opaque RGBA", NULL);
    engine_officeimo_free(pixels);
    CHECK("render tile", engine_officeimo_document_render_region(document, 1, 1.5, width / 2, height / 2, width - width / 2,
          height - height / 2, 0, &pixels, &pixel_length, &error), OFFICEIMO_OK);
    expect(pixel_length == (size_t)(width - width / 2) * (height - height / 2) * 4, "the tile has the wrong size", NULL);
    engine_officeimo_free(pixels);
    CHECK("render region off page", engine_officeimo_document_render_region(document, 1, 1.5, width, 0, 1, 1, 0, &pixels,
          &pixel_length, &error), OFFICEIMO_INVALID_ARGUMENT);

    CHECK("render missing page", engine_officeimo_document_render_page(document, 9999, 1.0, 0, &png, &length, NULL,
          NULL, &error), OFFICEIMO_INVALID_ARGUMENT);

    officeimo_cancellation cancellation = engine_officeimo_cancellation_create();
    engine_officeimo_cancellation_cancel(cancellation);
    CHECK("cancelled render", engine_officeimo_document_render_page(document, 1, 1.0, cancellation, &png, &length,
          NULL, NULL, &error), OFFICEIMO_CANCELLED);
    engine_officeimo_cancellation_free(cancellation);

    const int32_t first_page[] = { 1 };
    CHECK("rotate page 1", engine_officeimo_document_rotate_pages(document, first_page, 1, 90, 0, &error),
          OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "rotation", "90") && has(json, "isDirty", "true") && has(json, "canUndo", "true"),
           "rotation not recorded", json);
    expect(has(json, "canPrint", "true"), "an unrestricted document cannot be printed", json);
    engine_officeimo_free(json);

    CHECK("undo", engine_officeimo_document_undo(document, 0, &error), OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "rotation", "0") && has(json, "canRedo", "true"), "undo not applied", json);
    engine_officeimo_free(json);

    CHECK("redo", engine_officeimo_document_redo(document, 0, &error), OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "rotation", "90"), "redo not applied", json);
    engine_officeimo_free(json);

    /* Page organization: each edit is one undoable change, and the journey undoes them back to the rotated page. */
    CHECK("duplicate page 1", engine_officeimo_document_duplicate_pages(document, first_page, 1, 0, &error), OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "number", "2"), "the duplicate is missing", json);
    engine_officeimo_free(json);
    const int32_t reversed[] = { 2, 1 };
    CHECK("reorder pages", engine_officeimo_document_reorder_pages(document, reversed, 2, 0, &error), OFFICEIMO_OK);
    CHECK("reorder incomplete", engine_officeimo_document_reorder_pages(document, reversed, 1, 0, &error),
          OFFICEIMO_INVALID_ARGUMENT);
    CHECK("delete page 1", engine_officeimo_document_delete_pages(document, first_page, 1, 0, &error), OFFICEIMO_OK);
    CHECK("delete the last page", engine_officeimo_document_delete_pages(document, first_page, 1, 0, &error),
          OFFICEIMO_INVALID_ARGUMENT);
    CHECK("insert blank page", engine_officeimo_document_insert_blank_page(document, 2, 612, 792, 0, &error), OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "width", "612") && has(json, "height", "792"), "the blank page is missing", json);
    engine_officeimo_free(json);

    char extracted[1024];
    path_in(extracted, sizeof extracted, work, "extracted.pdf");
    CHECK("extract page 2", engine_officeimo_document_extract_pages(document, reversed, 1, extracted, 0, &error), OFFICEIMO_OK);
    officeimo_document extract = 0;
    CHECK("open extracted", engine_officeimo_document_open(extracted, NULL, 0, &extract, &error), OFFICEIMO_OK);
    if (extract) {
        json = facts(extract);
        expect(has(json, "width", "612") && !has(json, "number", "2"), "the extract is not the blank page", json);
        engine_officeimo_free(json);
        engine_officeimo_document_close(extract);
    }

    /* Import: the extracted page goes back in at the start as one change, then is undone. */
    char quoted[2200], import_request[2400];
    json_string(quoted, sizeof quoted, extracted);
    snprintf(import_request, sizeof import_request, "{\"insertBefore\":1,\"sources\":[{\"path\":%s}]}", quoted);
    CHECK("import pages", engine_officeimo_document_import_pages_json(document, import_request, 0, &json, &error), OFFICEIMO_OK);
    expect(has(json, "imported", "1"), "the page was not imported", json);
    engine_officeimo_free(json);
    CHECK("undo import", engine_officeimo_document_undo(document, 0, &error), OFFICEIMO_OK);
    CHECK("import nothing", engine_officeimo_document_import_pages_json(document, "{\"insertBefore\":1,\"sources\":[]}", 0, &json,
          &error), OFFICEIMO_INVALID_ARGUMENT);

    /* Make searchable: both pages, the rotated image and the blank one, gain the recognized word as text, which search
       then finds on each; undone after. */
    int recognized = 0;
    CHECK("make searchable", engine_officeimo_document_make_searchable_json(document, recognize_page, &recognized, 0, &json, &error),
          OFFICEIMO_OK);
    expect(recognized == 2 && has(json, "words", "2"), "the recognized words were not added", json);
    engine_officeimo_free(json);
    CHECK("search recognized", engine_officeimo_document_search_json(document, "journeyword", 0, &json, &error), OFFICEIMO_OK);
    expect(json && strstr(json, "\"page\":1") != NULL && strstr(json, "\"page\":2") != NULL, "the recognized words are not searchable",
           json);
    engine_officeimo_free(json);
    CHECK("undo searchable", engine_officeimo_document_undo(document, 0, &error), OFFICEIMO_OK);

    for (int edit = 0; edit < 4; edit++) CHECK("undo page edit", engine_officeimo_document_undo(document, 0, &error), OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "rotation", "90") && !has(json, "number", "2"), "page edits not undone", json);
    engine_officeimo_free(json);

    CHECK("page text", engine_officeimo_document_page_text_json(document, 1, 0, &json, &error), OFFICEIMO_OK);
    expect(json && strstr(json, "\"texts\":[") && strstr(json, "\"boxes\":[") && strstr(json, "\"breaks\":["),
           "the page text is missing", json);
    engine_officeimo_free(json);
    CHECK("page text missing page", engine_officeimo_document_page_text_json(document, 99, 0, &json, &error),
          OFFICEIMO_INVALID_ARGUMENT);
    CHECK("search", engine_officeimo_document_search_json(document, "no such words", 0, &json, &error), OFFICEIMO_OK);
    expect(has(json, "hits", "[]"), "unexpected search hits", json);
    engine_officeimo_free(json);

    /* Markup: an ink stroke and a note, listed, changed and removed, then undone back to the rotated page. */
    CHECK("add ink", engine_officeimo_document_add_annotation_json(document,
          "{\"tool\":\"ink\",\"page\":1,\"strokes\":[[{\"x\":20,\"y\":20},{\"x\":120,\"y\":60}]],"
          "\"color\":{\"red\":0.9,\"green\":0.1,\"blue\":0.1},\"strokeWidth\":4,\"opacity\":0.8}", 0, &error), OFFICEIMO_OK);
    CHECK("add note", engine_officeimo_document_add_annotation_json(document,
          "{\"tool\":\"note\",\"page\":1,\"bounds\":{\"x\":40,\"y\":40,\"width\":18,\"height\":18},\"text\":\"Journey\"}",
          0, &error), OFFICEIMO_OK);
    CHECK("add unknown tool", engine_officeimo_document_add_annotation_json(document, "{\"tool\":\"redact\",\"page\":1}", 0, &error),
          OFFICEIMO_INVALID_ARGUMENT);
    CHECK("list annotations", engine_officeimo_document_annotations_json(document, 1, 0, &json, &error), OFFICEIMO_OK);
    expect(has(json, "subtype", "\"Ink\"") && has(json, "contents", "\"Journey\""), "the markup is not listed", json);
    int note = 0;
    const char *found = json ? strstr(json, "\"subtype\":\"Text\"") : NULL;
    for (const char *at = found; at && at > json; at--) {
        if (strncmp(at, "\"objectNumber\":", 15) == 0) { note = atoi(at + 15); break; }
    }
    engine_officeimo_free(json);
    char update[128];
    snprintf(update, sizeof update, "{\"objectNumber\":%d,\"contents\":\"Checked\"}", note);
    CHECK("update note", engine_officeimo_document_update_annotation_json(document, update, 0, &error), OFFICEIMO_OK);
    CHECK("remove note", engine_officeimo_document_remove_annotation(document, note, 0, &error), OFFICEIMO_OK);
    CHECK("remove missing", engine_officeimo_document_remove_annotation(document, note, 0, &error), OFFICEIMO_INVALID_ARGUMENT);
    for (int edit = 0; edit < 4; edit++) CHECK("undo markup", engine_officeimo_document_undo(document, 0, &error), OFFICEIMO_OK);
    CHECK("list after undo", engine_officeimo_document_annotations_json(document, 1, 0, &json, &error), OFFICEIMO_OK);
    expect(has(json, "annotations", "[]"), "markup not undone", json);
    engine_officeimo_free(json);

    CHECK("save copy", engine_officeimo_document_save(document, copy, 0, &error), OFFICEIMO_OK);
    engine_officeimo_document_close(document);

    officeimo_document reopened = 0;
    CHECK("reopen copy", engine_officeimo_document_open(copy, NULL, 0, &reopened, &error), OFFICEIMO_OK);
    if (reopened) {
        json = facts(reopened);
        expect(has(json, "rotation", "90") && has(json, "isDirty", "false"), "the saved copy is not rotated", json);
        engine_officeimo_free(json);
        engine_officeimo_document_close(reopened);
    }

    /* A scan: the rendered page image becomes a PDF page with a recognized word as searchable text. */
    char scan[1024], scan_request[2048];
    path_in(scan, sizeof scan, work, "scan.pdf");
    snprintf(scan_request, sizeof scan_request,
             "{\"pages\":[{\"image\":\"%s\",\"words\":[{\"text\":\"Journey\",\"x\":0.1,\"y\":0.1,\"width\":0.3,\"height\":0.1}]}]}",
             page_png);
    CHECK("scan to pdf", engine_officeimo_create_pdf_from_scan_json(scan_request, scan, 0, &json, &error), OFFICEIMO_OK);
    expect(has(json, "searchablePages", "1"), "the scan is not searchable", json);
    engine_officeimo_free(json);
    officeimo_document scanned = 0;
    CHECK("open scan", engine_officeimo_document_open(scan, NULL, 0, &scanned, &error), OFFICEIMO_OK);
    if (scanned) {
        CHECK("search scan", engine_officeimo_document_search_json(scanned, "journey", 0, &json, &error), OFFICEIMO_OK);
        expect(has(json, "page", "1"), "the scan's word is not found", json);
        engine_officeimo_free(json);
        engine_officeimo_document_close(scanned);
    }
    CHECK("scan without pages", engine_officeimo_create_pdf_from_scan_json("{\"pages\":[]}", scan, 0, &json, &error),
          OFFICEIMO_INVALID_ARGUMENT);

    if (failures) {
        printf("%d step(s) failed.\n", failures);
        return 1;
    }
    printf("Native engine journey passed.\n");
    return 0;
}
