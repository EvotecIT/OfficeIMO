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

#define OFFICEIMO_ABI_VERSION 2

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

/*
 * Makes a PDF at output_path from scanned page images, one page per image at its size, as JSON:
 * { "pages": [ { "image": "<absolute path of a JPEG or PNG>", "words": [ { "text", "x", "y", "width", "height",
 *   "confidence" } ] } ] }
 * Words are what the host's recognizer found, each box from the image's top-left corner in fractions of the image's
 * width and height; they become invisible text over the image, so the scan can be searched, selected and copied. A page
 * without words is the image alone. The file is written beside output_path and replaces it only when complete. The
 * summary is JSON: { "abiVersion", "pages", "searchablePages", "words" }. No document needs to be open.
 */
int32_t officeimo_create_pdf_from_scan_json(const char *request, const char *output_path, officeimo_cancellation cancellation,
                                            char **json, char **error);

/* Cancellation: pass a token to any call; officeimo_cancellation_cancel may be called from any thread while it runs. */
officeimo_cancellation officeimo_cancellation_create(void);
void officeimo_cancellation_cancel(officeimo_cancellation cancellation);
void officeimo_cancellation_free(officeimo_cancellation cancellation);

/* Opens a PDF. password may be NULL; OFFICEIMO_PASSWORD_REQUIRED or OFFICEIMO_INVALID_PASSWORD asks the host to prompt. */
int32_t officeimo_document_open(const char *path, const char *password, officeimo_cancellation cancellation,
                                officeimo_document *document, char **error);

/* Opens an app-owned working file, keeping crash recovery under the original document's stable opaque identity.
 * identity is UTF-8, at most 4096 characters. NULL retains path-based recovery. No original file is read through it.
 * When info.hasRecovery is true, offer restore or discard before editing. Restore is undoable and marks the document
 * dirty; discard removes the pending snapshot without changing the document. Update identity when the original moves.
 */
int32_t officeimo_document_open_with_recovery_identity(const char *path, const char *password, const char *identity,
    officeimo_cancellation cancellation, officeimo_document *document, char **error);
int32_t officeimo_document_set_recovery_identity(officeimo_document document, const char *identity,
    officeimo_cancellation cancellation, char **error);
int32_t officeimo_document_restore_recovery(officeimo_document document, officeimo_cancellation cancellation, char **error);
int32_t officeimo_document_discard_recovery(officeimo_document document, officeimo_cancellation cancellation, char **error);
/*
 * Releases the document. Unsaved changes are discarded with their recovery copy; the host asks the user first. A host
 * that ends without closing keeps the recovery copy, which the next open reports as hasRecovery.
 */
void officeimo_document_close(officeimo_document document);

/*
 * Document facts as JSON:
 * { "abiVersion", "fileName", "path", "fileSize", "revision", "isDirty", "canUndo", "canRedo",
 *   "hasRecovery", "canExtractText", "canExtractContent", "pages": [ { "number", "width", "height", "rotation" } ],
 *   "maxRenderPixels", "canMutatePages", "canExtractPages", "canEditAnnotations", "canPrint", "canImportPages",
 *   "canMakeSearchable", "canFillForms", "isEncrypted", "canChangeProtection", "canSign", "canRotatePages",
 *   "canInsertBlankPages", "maxPrintDpi", "requiresOwnerPasswordForProtection" }
 * hasRecovery reports unsaved edits kept from an earlier session for this file. maxRenderPixels is the largest page
 * image officeimo_document_render_page produces. canExtractText allows search; canMutatePages allows deleting,
 * duplicating and reordering pages; canRotatePages and canInsertBlankPages describe their own operations independently.
 * canExtractPages allows officeimo_document_extract_pages; canEditAnnotations
 * allows adding, changing and removing annotations, which a tagged or signed document may take as an appended revision.
 * canPrint is false when the document's permissions forbid printing and it was not opened with its owner password;
 * maxPrintDpi is 150 for user-authorized low-resolution printing, otherwise null. Print hosts must honor this ceiling.
 * requiresOwnerPasswordForProtection is true only when owner authorization is the remaining protection-change barrier.
 * canImportPages allows officeimo_document_import_pages_json; canMakeSearchable allows officeimo_document_make_searchable_json;
 * canFillForms allows officeimo_document_fill_form_json. isEncrypted reports a password-protected document;
 * canChangeProtection allows officeimo_document_save_protected_copy_json without the owner password; canSign allows
 * officeimo_document_save_signed_copy_json.
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

/* Deletes pages as one undoable change. A document keeps at least one page. */
int32_t officeimo_document_delete_pages(officeimo_document document, const int32_t *pages, int32_t count,
                                        officeimo_cancellation cancellation, char **error);

/* Duplicates pages as one undoable change; each copy follows its page. */
int32_t officeimo_document_duplicate_pages(officeimo_document document, const int32_t *pages, int32_t count,
                                           officeimo_cancellation cancellation, char **error);

/* Puts the pages in a new order as one undoable change. order lists every current page number exactly once. */
int32_t officeimo_document_reorder_pages(officeimo_document document, const int32_t *order, int32_t count,
                                         officeimo_cancellation cancellation, char **error);

/*
 * Inserts a blank page of width by height points before page insert_before, or after the last page when insert_before
 * is the page count + 1, as one undoable change. Each side is between 3 and 14,400 points.
 */
int32_t officeimo_document_insert_blank_page(officeimo_document document, int32_t insert_before, double width, double height,
                                             officeimo_cancellation cancellation, char **error);

/* Writes pages to a new PDF at path, a different file from the document's; the open document is unchanged. */
int32_t officeimo_document_extract_pages(officeimo_document document, const int32_t *pages, int32_t count, const char *path,
                                         officeimo_cancellation cancellation, char **error);

/*
 * Inserts pages from other PDFs as one undoable change. The request is JSON:
 * { "insertBefore", "sources": [ { "path", "password", "pages": [ ] } ] }
 * Pages from each source, in the order listed, all of them when "pages" is absent or empty, go before page insertBefore,
 * or after the last page when it is the page count + 1; "password" opens a protected source. On success json receives
 * { "abiVersion", "imported", "revision" }. When a source needs its password, or the one given is wrong, the call fails
 * with OFFICEIMO_PASSWORD_REQUIRED or OFFICEIMO_INVALID_PASSWORD and json receives { "abiVersion", "source", "name" }:
 * the source's position in the request, from zero, and its file name, so the host can ask for that password and retry.
 */
int32_t officeimo_document_import_pages_json(officeimo_document document, const char *request,
                                             officeimo_cancellation cancellation, char **json, char **error);

/*
 * The form fields with a widget on a page, as JSON: { "abiVersion", "revision", "page", "canFill", "hasXfa", "fields": [ { "name",
 * "kind", "values": [ ], "options": [ { "value", "text" } ], "readOnly", "required", "multiline", "password", "comb",
 * "maxLength", "multipleSelection", "editable", "label", "widgets": [ { "objectNumber", "bounds": { "x", "y", "width",
 * "height" }, "onState", "isOn" } ] } ] }. kind is "text", "checkbox", "radio", "combo", "list", "button" or "signature";
 * bounds are in points of the page as displayed; onState is the value that turns a check box or radio button on.
 * hasXfa reports whether the form also carries an XFA packet.
 */
int32_t officeimo_document_form_fields_json(officeimo_document document, int32_t page, officeimo_cancellation cancellation,
                                            char **json, char **error);

/*
 * Sets form field values as one undoable change. The request is JSON:
 * { "values": { "<field name>": [ "<value>" ] }, "removeXfa": false }:
 * text, the selected option values of a choice, or a widget's onState for a check box or radio button, "Off" to clear
 * it. A read-only, push-button or signature field, or a name the form doesn't have, fails with OFFICEIMO_INVALID_ARGUMENT.
 * removeXfa defaults to false, preserving the XFA packet without synchronizing its data. Set it to true to remove
 * XFA in the appended revision so XFA-aware viewers use the AcroForm values; this can invalidate Reader usage rights.
 * Removing XFA from a DocMDP-certified form is refused. Fill requests must preserve XFA and satisfy certification permissions.
 */
int32_t officeimo_document_fill_form_json(officeimo_document document, const char *request, officeimo_cancellation cancellation,
                                          char **error);

/*
 * The document formats that convert to PDF, as JSON: { "abiVersion", "sources": [ { "route", "name", "extensions": [ ] } ] },
 * such as Word, Excel and PowerPoint files, with the file extensions each is read from, ".docx" and so on.
 */
int32_t officeimo_conversion_sources_json(char **json, char **error);

/*
 * Converts the document at input_path to a PDF at output_path, replacing what is there, through the route its extension
 * names in officeimo_conversion_sources_json. json receives { "abiVersion", "route", "source", "outputBytes",
 * "notes": [ { "severity", "message" } ] }: the converter's warnings about content the PDF shows differently. A format
 * with no route fails with OFFICEIMO_UNSUPPORTED; a document the converter can't read fails with its reason.
 */
int32_t officeimo_convert_to_pdf_json(const char *input_path, const char *output_path, officeimo_cancellation cancellation,
                                      char **json, char **error);

/*
 * Recognizes a page's text for officeimo_document_make_searchable_json. It receives the context given to that call, the
 * page number from 1, the page count and the page rendered as PNG, and sets *words_json to a JSON array allocated with
 * malloc, which the engine frees: [ { "text", "x", "y", "width", "height", "confidence" } ], each box from the image's
 * top-left corner in fractions of its width and height, and confidence from 0 to 1 or null. It returns OFFICEIMO_OK,
 * OFFICEIMO_CANCELLED to stop, or another status to fail the call. It runs on the thread that made the call.
 */
typedef int32_t (*officeimo_text_recognizer)(void *context, int32_t page, int32_t page_count, const uint8_t *png, size_t length,
                                             char **words_json);

/*
 * Makes the document searchable as one undoable change: each page is rendered and passed to recognizer, and the words it
 * returns become invisible text over the page, so the document can be searched, selected and copied. Words over text a
 * page already has are left out, so digital pages are unchanged. json receives { "abiVersion", "pages", "words",
 * "revision" }: the pages that gained text and the words added.
 */
int32_t officeimo_document_make_searchable_json(officeimo_document document, officeimo_text_recognizer recognizer, void *context,
                                                officeimo_cancellation cancellation, char **json, char **error);

/*
 * Finds text, ignoring case, as JSON:
 * { "abiVersion", "query", "revision", "hits": [ { "page", "occurrence", "text",
 *   "bounds": { "x", "y", "width", "height" }, "lineBounds": [ { "x", "y", "width", "height" } ] } ] }
 * Rectangles are in points of the page as displayed, from its top-left corner after rotation: the page's width and
 * height, swapped for 90 and 270 degrees. lineBounds has one rectangle per line a hit spans; bounds is their union.
 * Occurrences are numbered from 1 in document order. A document that restricts text extraction fails with
 * OFFICEIMO_UNSUPPORTED.
 */
int32_t officeimo_document_search_json(officeimo_document document, const char *query, officeimo_cancellation cancellation,
                                       char **json, char **error);

/*
 * Receives one page's search results while officeimo_document_search_stream_json runs. It receives the context given
 * to that call and UTF-8 JSON valid only during the call:
 * { "abiVersion", "query", "revision", "pagesSearched", "pageCount", "hits": [ ... ] }
 * hits are the hits on the page just searched, as officeimo_document_search_json describes them, numbered in document
 * order; a page without hits reports an empty array. It runs on the engine's thread, once per page in order.
 */
typedef void (*officeimo_search_progress)(void *context, const char *page_json);

/*
 * Finds text as officeimo_document_search_json does, passing each page's hits to progress as soon as the page is
 * searched, so a host shows hits and progress while a long document is searched. Cancelling stops between pages. The
 * returned JSON is the complete result.
 */
int32_t officeimo_document_search_stream_json(officeimo_document document, const char *query, officeimo_search_progress progress,
                                              void *context, officeimo_cancellation cancellation, char **json, char **error);

/*
 * Adds a markup annotation as one undoable change, described as JSON:
 * { "tool", "page", "bounds": { "x", "y", "width", "height" }, "strokes": [ [ { "x", "y" } ] ],
 *   "quads": [ { "x", "y", "width", "height" } ], "text", "author", "color": { "red", "green", "blue" },
 *   "strokeWidth", "opacity", "fontSize", "linkUri", "linkPage", "stamp", "lineEnd" }
 * tool is one of ink, note, freeText, highlight, underline, strikeOut, squiggly, rectangle, ellipse, line, polygon, link
 * and stamp. A link covers its bounds and goes to linkUri (https, http or mailto) or to linkPage, exactly one of them; a
 * stamp is one of the standard stamps, such as Approved, Draft, Confidential, Final or ForComment (default Draft). A
 * line's lineEnd is None, OpenArrow or ClosedArrow, drawn at its last point.
 * Geometry is in points of the page as displayed, from its top-left corner, the space of officeimo_document_search_json.
 * Ink takes its strokes; a line its first stroke's first and last points; a polygon its first stroke; text markup its
 * quads, or bounds without them; the others take bounds. bounds may be left out when strokes are given. Colours are
 * from 0 through 1, strokeWidth is in points and opacity above 0 through 1; text is the annotation's contents.
 */
/* request also accepts an array of these annotations, committed atomically as one undoable revision. */
int32_t officeimo_document_add_annotation_json(officeimo_document document, const char *request,
                                               officeimo_cancellation cancellation, char **error);

/*
 * The markup annotations on a page as JSON:
 * { "abiVersion", "revision", "page", "canEdit", "annotations": [ { "objectNumber", "page", "subtype",
 *   "bounds": { "x", "y", "width", "height" }, "contents", "author", "color": { "red", "green", "blue" } } ] }
 * Bounds are in points of the page as displayed. Links are listed; form fields and popups are not.
 */
int32_t officeimo_document_annotations_json(officeimo_document document, int32_t page, officeimo_cancellation cancellation,
                                            char **json, char **error);

/*
 * Changes an annotation's contents, author or colour as one undoable change: { "objectNumber", "contents", "author",
 * "color": { "red", "green", "blue" } }. Fields left out are kept. Alone, "bounds": { "x", "y", "width", "height" },
 * in points of the page as displayed, moves the annotation there, or resizes it with its geometry when the size
 * differs.
 */
int32_t officeimo_document_update_annotation_json(officeimo_document document, const char *update,
                                                  officeimo_cancellation cancellation, char **error);

/* Removes an annotation, by its objectNumber from officeimo_document_annotations_json, as one undoable change. */
int32_t officeimo_document_remove_annotation(officeimo_document document, int32_t object_number,
                                             officeimo_cancellation cancellation, char **error);

/*
 * A page's text for selection and copying, as JSON:
 * { "abiVersion", "revision", "page", "texts": [ ], "boxes": [ ], "breaks": [ ] }.
 * texts holds the page's text elements, usually one character each, in extraction order; boxes holds four numbers per
 * element, x, y, width and height, in points of the page as displayed from its top-left corner. breaks holds one number
 * per element for how it follows the previous one: 0 continues it, 1 is further along the same line after a gap, as
 * between separately positioned words of a scan's text, and 2 starts another line. A document that restricts text
 * extraction fails with OFFICEIMO_UNSUPPORTED.
 */
int32_t officeimo_document_page_text_json(officeimo_document document, int32_t page, officeimo_cancellation cancellation,
                                          char **json, char **error);

int32_t officeimo_document_undo(officeimo_document document, officeimo_cancellation cancellation, char **error);
int32_t officeimo_document_redo(officeimo_document document, officeimo_cancellation cancellation, char **error);

/*
 * Saves in place when path is NULL, otherwise to path, which becomes the document's location. The output is written
 * beside the destination and replaces it only when complete, so a failed save leaves the previous file intact. Saving in
 * place fails if the file was replaced or moved since it was opened; save to another location instead.
 */
int32_t officeimo_document_save(officeimo_document document, const char *path, officeimo_cancellation cancellation,
                                char **error);

/*
 * Writes the open revision, including unsaved changes, to path, as a share or an export needs. The document keeps its
 * location and saved state. path must not be the document's own file.
 */
int32_t officeimo_document_save_copy(officeimo_document document, const char *path, officeimo_cancellation cancellation,
                                     char **error);

/* Captures immutable bytes for a host-owned save; release *pdf with officeimo_free. Recovery remains until acknowledged. */
int32_t officeimo_document_prepare_save(officeimo_document document, officeimo_cancellation cancellation,
                                        uint8_t **pdf, size_t *length, char **error);
/* Acknowledge only a prepared SHA-256 actually observed at the host's original file, after its publication completes. */
int32_t officeimo_document_acknowledge_save(officeimo_document document, const char *fingerprint, const char *identity,
                                            officeimo_cancellation cancellation, char **error);

/*
 * Writes the open revision with a password to open it, described as JSON:
 * { "path", "userPassword", "ownerPassword", "currentOwnerPassword", "allowPrinting", "allowCopying", "allowChanges" }
 * ownerPassword unlocks the permissions and is required when any allow flag is false. It must differ from userPassword
 * after AES-256 FormKC normalization and truncation to 127 UTF-8 bytes. The allow flags default to true. An empty userPassword writes the
 * document without password protection. allowCopying=false disables both content copying and accessibility text extraction.
 * currentOwnerPassword is needed when the document was opened with its
 * user password and its permissions forbid changing security. The document itself keeps its protection and location.
 */
int32_t officeimo_document_save_protected_copy_json(officeimo_document document, const char *request,
                                                    officeimo_cancellation cancellation, char **error);

/*
 * Writes the open revision with an invisible certificate signature, described as JSON:
 * { "path", "certificatePath", "certificatePassword", "reason", "location" }
 * certificatePath is a PKCS #12 (.p12 or .pfx) file with the signer's certificate and private key; a password it doesn't
 * accept fails with OFFICEIMO_INVALID_PASSWORD. The signature covers the whole copy, so a reader can verify who signed it
 * and that it hasn't changed since. The document itself is unchanged.
 */
int32_t officeimo_document_save_signed_copy_json(officeimo_document document, const char *request, officeimo_cancellation cancellation,
                                                 char **error);

/*
 * The document's structure as JSON: { "abiVersion", "revision", "properties": { "title", "author", "subject", "keywords",
 * "created", "modified" }, "canEditProperties", "outline": [ { "id", "title", "page", "children": [ ] } ], "canEditOutline",
 * "attachments": [ { "id", "fileName", "size", "description", "mimeType" } ], "canEditAttachments", "canCrop", "canResize",
 * "canStampPages", "hasForms", "canFlattenForms", "canFlattenMarkup" }. created and modified are ISO 8601 times. An outline entry without an id
 * has no page in this document; it and its children are shown but can't be changed.
 */
int32_t officeimo_document_structure_json(officeimo_document document, officeimo_cancellation cancellation, char **json,
                                          char **error);

/*
 * Applies one edit as one undoable change, described as JSON naming exactly one of:
 *   "properties": { "title", "author", "subject", "keywords" }: a field left out keeps its value; "" clears it.
 *   "outline": [ { "op", "id", "title", "page", "parentId", "index" } ]: op is "add" (title, page, optional parentId and
 *     index), "rename" (id, title), "remove" (id), "move" (id, parentId, index), "retarget" (id, page) or "rebuild", which
 *     makes the outline from the document's headings. Ids come from officeimo_document_structure_json; all the changes in
 *     one request are one undoable step.
 *   "crop": { "pages", "left", "top", "right", "bottom" }: margins in points cut from each edge of the pages as displayed.
 *   "pageSize": { "pages", "width", "height", "mode" }: a new paper size in points, mode "fit" (default), "fill" or "stretch".
 *   "headerFooter": { "headerLeft", "headerCenter", "headerRight", "footerLeft", "footerCenter", "footerRight", "fontSize",
 *     "margin", "startNumber", "pages" }: text accepts {page}, {pages}, {date}, {file} and {title}.
 *   "watermark": { "text", "fontSize", "opacity", "rotation", "behindContent", "color": { "red", "green", "blue" }, "pages" }.
 *   "bates": { "prefix", "suffix", "start", "digits", "position" }: position is "bottomRight" (default), "bottomCenter",
 *     "bottomLeft", "topLeft", "topCenter" or "topRight"; every page is numbered.
 *   "flatten": { "fields" }: draws the named field, or every field when none is named, into its page; or
 *     { "markup": true }: draws every markup annotation, but links, into its page.
 *   "addAttachment": { "path", "fileName" }, or "removeAttachment": "file name".
 * Leaving out "pages" means every page. The structure's can flags tell which edits the document allows.
 */
int32_t officeimo_document_edit_json(officeimo_document document, const char *request, officeimo_cancellation cancellation,
                                     char **error);

/*
 * Writes an output, leaving the document unchanged, described as JSON naming exactly one of:
 *   "split": { "directory", "pagesPerPart" } or { "directory", "parts": [ { "pages", "name" } ] }: at least two parts,
 *     each a new PDF in directory, named from the document and its pages unless a name is given.
 *   "attachment": { "id", "path" }: saves the identified attachment's contents. A unique fileName is also accepted.
 *   "export": { "format", "path" }: the document's text as "text", "markdown" or structured "json"; the document converted
 *     to "word", "excel" (its tables), "powerpoint" or "html"; or, with "images", its embedded images saved into the folder
 *     path names.
 *   "optimize": { "path", "profile" }: a losslessly smaller copy; profile is "balanced" (default), "smallest", "web" or
 *     "archival". The copy is written only when it keeps everything the document contains.
 *   "repair": { "path" }: a copy with the structural problems the reader recovered written out, verified by reopening it.
 *   "sanitize": { "path" }: a copy without scripts, actions, attachments and other active content, verified afterwards.
 * Never replaces an existing part: a name already taken gets a number. A health copy replaces the file at path, which
 * must not be the document's own. Returns { "abiVersion", "revision", "files", "summary" }, summary saying what a
 * health copy changed or, one per line, what a conversion couldn't keep.
 */
int32_t officeimo_document_write_json(officeimo_document document, const char *request, officeimo_cancellation cancellation,
                                      char **json, char **error);

/*
 * Checks the document without changing it, described as JSON naming exactly one of:
 *   "signatures": true -> "signatures": { "hasSignatures", "signatures": [ { "signer", "field", "signingTime", "reason",
 *     "location", "status", "coversWholeDocument", "certificate", "timestamp", "findings": [ ] } ], "findings": [ ] }.
 *     status is "valid" (intact and from a trusted certificate), "intact" (unchanged since signing, certificate not
 *     trusted), "invalid", "unknown" or "unchecked"; certificate and timestamp are "valid", "invalid", "unknown" or
 *     "notChecked".
 *   "standard": "pdfa2b", "pdfa2u", "pdfa2a", "pdfa3b", "pdfa3u", "pdfa3a", "pdfa4", "pdfua1", "pdfua2", "pdfx1a" or "pdfx4"
 *     -> "standard": { "name", "isReady", "requirements": [ { "id", "name", "status", "detail" } ] }, status "met",
 *     "missing" or "unsupported".
 *   "health": true -> "health": { "fileSize", "pages", "version", "canRead", "canRewrite", "isEncrypted", "isSigned",
 *     "isTagged", "hasActiveContent", "hasAttachments", "recoveredProblems", "unrecoverableProblems", "diagnostics" }.
 * Returns { "abiVersion", "revision" } with the one report.
 */
int32_t officeimo_document_check_json(officeimo_document document, const char *request, officeimo_cancellation cancellation,
                                      char **json, char **error);

/*
 * Redaction, described as JSON naming exactly one of:
 *   "find": { "text", "regex", "matchCase", "pages" } -> { "marks": [ { "page", "bounds", "preview" } ] }: the found text's
 *     own glyphs, in points of the page as displayed; the document is unchanged.
 *   "apply": { "revision", "areas": [ { "page", "bounds" } ], "find", "include", "removedText", "sanitize" } -> { "verification":
 *     { "verified", "removed", "residual", "inconclusive", "pages" } }: removes the text, images and drawings under each
 *     drawn area and each match of find (all of them, or the marks listed by index in include), paints them black and
 *     checks that the content is gone, as one undoable change. find is searched again on the same revision, so it keeps
 *     the exact glyphs. revision is required: use the find result's revision, or the document facts revision on which
 *     drawn areas were reviewed. A missing or stale revision is rejected before mutation, including after undo/redo.
 *     removedText, when given, must be absent afterwards; sanitize also removes metadata,
 *     attachments, actions and markup.
 * Returns { "abiVersion", "revision" } with marks or verification.
 */
int32_t officeimo_document_redact_json(officeimo_document document, const char *request, officeimo_cancellation cancellation,
                                       char **json, char **error);

#ifdef __cplusplus
}
#endif

#endif /* OFFICEIMO_STUDIO_H */
