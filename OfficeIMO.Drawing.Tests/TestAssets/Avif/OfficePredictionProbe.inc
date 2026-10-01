/* Opt-in access to actual native prediction builders; their algorithms are unchanged. */
void office_intra_prediction(const int *p, const uint8_t *above, const uint8_t *left, uint8_t *output) {
  uint8_t frame[256 * 132];
  memset(frame, 0, sizeof(frame));
  uint8_t *ref = frame + 256 + 1;
  ref[-257] = (uint8_t)p[10];
  for (int i = 0; i < p[11]; i++) ref[-256 + i] = above[i];
  for (int i = 0; i < p[12]; i++) ref[i * 256 - 1] = left[i];
  av1_init_intra_predictors();
  if (p[5] >= 0 || av1_is_directional_mode((PREDICTION_MODE)p[1]))
    build_directional_and_filter_intra_predictors(ref, 256, output, tx_size_wide[p[0]],
      (PREDICTION_MODE)p[1], mode_to_angle_map[p[1]] + 3 * p[2],
      p[5] < 0 ? FILTER_INTRA_MODES : (FILTER_INTRA_MODE)p[5], (TX_SIZE)p[0], !p[3],
      p[6], p[8], p[7], p[9], p[4]);
  else build_non_directional_intra_predictors(ref, 256, output, tx_size_wide[p[0]],
      (PREDICTION_MODE)p[1], (TX_SIZE)p[0], p[6], p[7]);
}
void office_palette_prediction(const int *p, const uint8_t *colors, const uint8_t *map, uint8_t *output) {
  MACROBLOCKD xd; MB_MODE_INFO mbmi; YV12_BUFFER_CONFIG buffer;
  memset(&xd, 0, sizeof(xd)); memset(&mbmi, 0, sizeof(mbmi)); memset(&buffer, 0, sizeof(buffer));
  MB_MODE_INFO *mi = &mbmi; xd.mi = &mi; xd.cur_buf = &buffer;
  xd.plane[p[1] != 0].color_index_map = map;
  for (int i = 0; i < p[6]; i++) mbmi.palette_mode_info.palette_colors[p[1] * PALETTE_MAX_SIZE + i] = colors[i];
  av1_predict_intra_block(&xd, BLOCK_64X64, 0, p[2], p[3], (TX_SIZE)p[0], DC_PRED,
      0, 1, FILTER_INTRA_MODES, NULL, 0, output, tx_size_wide[p[0]], p[4] / 4, p[5] / 4, p[1]);
}
