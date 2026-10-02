/* Observation-only entry point appended to the pinned native decodemv.c.
 * Palette syntax, cache selection, CDF updates and index decoding stay native. */
#include <stdio.h>
#include <stdlib.h>
#include "av1/decoder/detokenize.h"

typedef struct {
  AV1_COMMON cm;
  SequenceHeader sequence;
  MACROBLOCKD xd;
  FRAME_CONTEXT context;
  MB_MODE_INFO current, *current_pointer, grid[64 * 64];
  aom_reader reader;
  uint8_t y_map[4096], uv_map[1024];
  int extent;
} OfficePaletteProbe;

void *office_palette_create(const uint8_t *bytes, int length, int depth,
                           int extent, int updates, int screen, int filter,
                           int mono) {
  OfficePaletteProbe *p = calloc(1, sizeof(*p));
  if (!p || aom_reader_init(&p->reader, bytes, length)) abort();
  p->extent = extent;
  p->sequence.bit_depth = (aom_bit_depth_t)depth;
  p->sequence.monochrome = mono;
  p->sequence.enable_filter_intra = filter;
  p->cm.seq_params = &p->sequence;
  p->cm.features.allow_screen_content_tools = screen;
  av1_init_mode_probs(&p->context);
  p->reader.allow_update_cdf = updates;
  p->current_pointer = &p->current;
  p->xd.mi = &p->current_pointer;
  p->xd.tile_ctx = &p->context;
  p->xd.plane[0].color_index_map = p->y_map;
  p->xd.plane[1].color_index_map = p->uv_map;
  p->xd.plane[1].subsampling_x = p->xd.plane[1].subsampling_y = 1;
  return p;
}

static void office_palette_colors(FILE *out, const uint16_t *colors, int n) {
  fputc('[', out);
  for (int i = 0; i < n; ++i) fprintf(out, "%s%u", i ? "," : "", colors[i]);
  fputc(']', out);
}
static void office_palette_map(FILE *out, const uint8_t *map, int n) {
  fputc('"', out);
  for (int i = 0; i < n; ++i) fprintf(out, "%02x", map[i]);
  fputc('"', out);
}

void office_palette_leaf(void *handle, const int *v, FILE *out) {
  OfficePaletteProbe *p = handle;
  MACROBLOCKD *xd = &p->xd;
  const int row = v[0], col = v[1], width = v[2], height = v[3];
  memset(&p->current, 0, sizeof(p->current));
  BLOCK_SIZE size = BLOCK_INVALID;
  for (int i = 0; i < BLOCK_SIZES_ALL; ++i)
    if (block_size_wide[i] == width && block_size_high[i] == height) size = i;
  if (size == BLOCK_INVALID || row < 0 || col < 0 || row >= 64 || col >= 64)
    abort();
  p->current.bsize = size;
  p->current.mode = (PREDICTION_MODE)v[7];
  p->current.uv_mode = (UV_PREDICTION_MODE)v[8];
  p->current.use_intrabc = v[4];
  xd->is_chroma_ref = v[5];
  xd->above_mbmi = row ? &p->grid[(row - 1) * 64 + col] : NULL;
  xd->left_mbmi = col ? &p->grid[row * 64 + col - 1] : NULL;
  xd->up_available = row != 0;
  xd->left_available = col != 0;
  xd->mb_to_top_edge = -row * 4 * 8;
  xd->mb_to_right_edge = (p->extent - col - width / 4) * 4 * 8;
  xd->mb_to_bottom_edge = (p->extent - row - height / 4) * 4 * 8;
  if (!p->current.use_intrabc) {
    if (av1_allow_palette(p->cm.features.allow_screen_content_tools, size))
      read_palette_mode_info(&p->cm, xd, &p->reader);
    read_filter_intra_mode_info(&p->cm, xd, &p->reader);
  }
  PALETTE_MODE_INFO *palette = &p->current.palette_mode_info;
  if (palette->palette_size[0]) av1_decode_palette_tokens(xd, 0, &p->reader);
  if (palette->palette_size[1]) av1_decode_palette_tokens(xd, 1, &p->reader);
  fprintf(out, "{\"filter\":%d,\"y\":",
          p->current.filter_intra_mode_info.use_filter_intra
              ? p->current.filter_intra_mode_info.filter_intra_mode : -1);
  office_palette_colors(out, palette->palette_colors, palette->palette_size[0]);
  fputs(",\"u\":", out);
  office_palette_colors(out, palette->palette_colors + 8, palette->palette_size[1]);
  fputs(",\"v\":", out);
  office_palette_colors(out, palette->palette_colors + 16, palette->palette_size[1]);
  fputs(",\"mapY\":", out);
  office_palette_map(out, p->y_map, palette->palette_size[0] ? width * height : 0);
  fputs(",\"mapUv\":", out);
  office_palette_map(out, p->uv_map, palette->palette_size[1]
      ? AOMMAX(4, width / 2) * AOMMAX(4, height / 2) : 0);
  fputc('}', out);
  for (int r = row; r < AOMMIN(p->extent, row + height / 4); ++r)
    for (int c = col; c < AOMMIN(p->extent, col + width / 4); ++c)
      p->grid[r * 64 + c] = p->current;
}

void office_palette_destroy(void *handle) {
  OfficePaletteProbe *p = handle;
  if (aom_reader_has_overflowed(&p->reader)) abort();
  free(p);
}
