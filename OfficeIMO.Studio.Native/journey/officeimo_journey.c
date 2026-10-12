/*
 * The engine's document journey from a native host: open, read facts, render, rotate, undo, redo, save a copy and
 * reopen it. A .NET process cannot host the library on macOS or Linux, because its runtime and the library's handle
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
    X(officeimo_document_render_page) X(officeimo_document_rotate_pages) X(officeimo_document_undo) \
    X(officeimo_document_redo) X(officeimo_document_save)

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

static void expect(int condition, const char *description, const char *json) {
    if (!condition) {
        printf("FAILED: %s: %s\n", description, json ? json : "(no facts)");
        failures++;
    }
}

/* Reads the document facts. The caller releases the JSON with officeimo_free. */
static char *facts(officeimo_document document) {
    char *json = NULL, *error = NULL;
    check("info", engine_officeimo_document_info_json(document, &json, &error), error, OFFICEIMO_OK);
    return json;
}

/* Finds "name": followed by value in the facts. The journey uses one-page documents, so page fields are unique. */
static int has(const char *json, const char *name, const char *value) {
    char pattern[128];
    snprintf(pattern, sizeof pattern, "\"%s\":%s", name, value);
    return json && strstr(json, pattern) != NULL;
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
    check("open before initialize", engine_officeimo_document_open(source, NULL, 0, &document, &error), error,
          OFFICEIMO_NOT_INITIALIZED);
    check("initialize", engine_officeimo_initialize(data, &error), error, OFFICEIMO_OK);

    double started = milliseconds();
    check("open", engine_officeimo_document_open(source, NULL, 0, &document, &error), error, OFFICEIMO_OK);
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
    check("render page 1", engine_officeimo_document_render_page(document, 1, 1.5, 0, &png, &length, &width, &height,
          &error), error, OFFICEIMO_OK);
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
    check("render missing page", engine_officeimo_document_render_page(document, 9999, 1.0, 0, &png, &length, NULL,
          NULL, &error), error, OFFICEIMO_INVALID_ARGUMENT);

    officeimo_cancellation cancellation = engine_officeimo_cancellation_create();
    engine_officeimo_cancellation_cancel(cancellation);
    check("cancelled render", engine_officeimo_document_render_page(document, 1, 1.0, cancellation, &png, &length,
          NULL, NULL, &error), error, OFFICEIMO_CANCELLED);
    engine_officeimo_cancellation_free(cancellation);

    const int32_t first_page[] = { 1 };
    check("rotate page 1", engine_officeimo_document_rotate_pages(document, first_page, 1, 90, 0, &error), error,
          OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "rotation", "90") && has(json, "isDirty", "true") && has(json, "canUndo", "true"),
           "rotation not recorded", json);
    engine_officeimo_free(json);

    check("undo", engine_officeimo_document_undo(document, 0, &error), error, OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "rotation", "0") && has(json, "canRedo", "true"), "undo not applied", json);
    engine_officeimo_free(json);

    check("redo", engine_officeimo_document_redo(document, 0, &error), error, OFFICEIMO_OK);
    json = facts(document);
    expect(has(json, "rotation", "90"), "redo not applied", json);
    engine_officeimo_free(json);

    check("save copy", engine_officeimo_document_save(document, copy, 0, &error), error, OFFICEIMO_OK);
    engine_officeimo_document_close(document);

    officeimo_document reopened = 0;
    check("reopen copy", engine_officeimo_document_open(copy, NULL, 0, &reopened, &error), error, OFFICEIMO_OK);
    if (reopened) {
        json = facts(reopened);
        expect(has(json, "rotation", "90") && has(json, "isDirty", "false"), "the saved copy is not rotated", json);
        engine_officeimo_free(json);
        engine_officeimo_document_close(reopened);
    }

    if (failures) {
        printf("%d step(s) failed.\n", failures);
        return 1;
    }
    printf("Native engine journey passed.\n");
    return 0;
}
