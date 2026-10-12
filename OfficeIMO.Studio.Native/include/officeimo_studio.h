/*
 * OfficeIMO Studio engine: C interface for native host applications (macOS, iPadOS, iOS).
 *
 * Conventions
 * - Strings are UTF-8 and NUL-terminated. Pages are numbered from 1.
 * - Every call blocks until the engine finishes. Call from a background thread or task, never the main thread.
 *   Calls on one document are serialized by the engine; different documents may be used concurrently.
 * - Fallible calls return an officeimo_status. On failure, *error (when error is not NULL) receives a message the
 *   host may show; on success it is set to NULL.
 * - Memory returned by the library (strings, JSON, image bytes, error messages) is released with officeimo_free.
 *   Handles have their own release functions and must be released exactly once.
 * - Paths are local file-system paths. On Apple platforms the host resolves security-scoped URLs and keeps access
 *   open for the duration of each call that reads or writes the file; the engine never retains that access.
 * - officeimo_abi_version() changes whenever a signature or JSON payload changes incompatibly. Hosts check it at launch.
 */
#ifndef OFFICEIMO_STUDIO_H
#define OFFICEIMO_STUDIO_H

#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

#define OFFICEIMO_ABI_VERSION 1

typedef enum officeimo_status {
    OFFICEIMO_OK = 0,
    OFFICEIMO_FAILED = 1,
    OFFICEIMO_CANCELLED = 2,
    OFFICEIMO_INVALID_ARGUMENT = 3,
    OFFICEIMO_PASSWORD_REQUIRED = 4,
    OFFICEIMO_INVALID_PASSWORD = 5,
    OFFICEIMO_UNSUPPORTED = 6,
    OFFICEIMO_NOT_FOUND = 7,
    OFFICEIMO_NOT_INITIALIZED = 8
} officeimo_status;

/* Opaque handles. 0 is never a valid handle. */
typedef intptr_t officeimo_document;
typedef intptr_t officeimo_cancellation;

int32_t officeimo_abi_version(void);
/* The engine version, released with officeimo_free. */
char *officeimo_library_version(void);
void officeimo_free(void *memory);

/* Names the absolute, app-owned folder (inside the app container) for recovery copies. Call once before opening documents. */
int32_t officeimo_initialize(const char *data_directory, char **error);

/* Cancellation: pass a token to any call; officeimo_cancellation_cancel may be called from any thread while it runs. */
officeimo_cancellation officeimo_cancellation_create(void);
void officeimo_cancellation_cancel(officeimo_cancellation cancellation);
void officeimo_cancellation_free(officeimo_cancellation cancellation);

/* Opens a PDF. password may be NULL; OFFICEIMO_PASSWORD_REQUIRED or OFFICEIMO_INVALID_PASSWORD asks the host to prompt. */
int32_t officeimo_document_open(const char *path, const char *password, officeimo_cancellation cancellation,
                                officeimo_document *document, char **error);
/*
 * Releases the document. Unsaved changes are discarded with their recovery copy; the host asks the user first. A host
 * that ends without closing keeps the recovery copy, which the next open reports as hasRecovery.
 */
void officeimo_document_close(officeimo_document document);

/*
 * Document facts as JSON:
 * { "abiVersion", "fileName", "path", "fileSize", "revision", "isDirty", "canUndo", "canRedo",
 *   "hasRecovery", "canExtractText", "canExtractContent", "pages": [ { "number", "width", "height", "rotation" } ],
 *   "maxRenderPixels" }
 * hasRecovery reports unsaved edits kept from an earlier session for this file. maxRenderPixels is the largest page
 * image officeimo_document_render_page produces.
 * Width and height are in points (1/72 inch) before rotation: the page's visible box at its user unit, the size that
 * render scales apply to. The revision changes with every edit, undo and redo.
 */
int32_t officeimo_document_info_json(officeimo_document document, char **json, char **error);

/*
 * Renders a page as PNG bytes at scale pixels per point, for thumbnails, sharing and export; display uses
 * officeimo_document_render_region. pixel_width and pixel_height may be NULL. A scale whose page
 * image would exceed maxRenderPixels fails with OFFICEIMO_INVALID_ARGUMENT; the host renders a smaller scale.
 */
int32_t officeimo_document_render_page(officeimo_document document, int32_t page, double scale,
                                       officeimo_cancellation cancellation, uint8_t **png, size_t *length,
                                       int32_t *pixel_width, int32_t *pixel_height, char **error);

/*
 * Renders a rectangle of a page as pixels for display. The page image at scale pixels per point is the page's width and
 * height times scale, each rounded up; x, y, width and height select a rectangle of it, all of it or a tile. Tiles of
 * one page at one scale join without seams, so zooming past maxRenderPixels renders the visible tiles. width * height
 * must not exceed maxRenderPixels. The pixels are opaque RGBA, 8 bits per channel, rows top to bottom, 4 * width bytes
 * per row: *length is 4 * width * height. The engine keeps recently rendered pages and their decoded images, so further
 * tiles, zoom levels and repeated views of a page are faster until the document changes.
 */
int32_t officeimo_document_render_region(officeimo_document document, int32_t page, double scale, int32_t x, int32_t y,
                                         int32_t width, int32_t height, officeimo_cancellation cancellation,
                                         uint8_t **pixels, size_t *length, char **error);

/* Rotates pages by a multiple of 90 degrees as one undoable change. */
int32_t officeimo_document_rotate_pages(officeimo_document document, const int32_t *pages, int32_t count,
                                        int32_t degrees, officeimo_cancellation cancellation, char **error);
int32_t officeimo_document_undo(officeimo_document document, officeimo_cancellation cancellation, char **error);
int32_t officeimo_document_redo(officeimo_document document, officeimo_cancellation cancellation, char **error);

/*
 * Saves in place when path is NULL, otherwise to path, which becomes the document's location. The output is written
 * beside the destination and replaces it only when complete, so a failed save leaves the previous file intact. Saving in
 * place fails if the file was replaced or moved since it was opened; save to another location instead.
 */
int32_t officeimo_document_save(officeimo_document document, const char *path, officeimo_cancellation cancellation,
                                char **error);

#ifdef __cplusplus
}
#endif

#endif /* OFFICEIMO_STUDIO_H */
