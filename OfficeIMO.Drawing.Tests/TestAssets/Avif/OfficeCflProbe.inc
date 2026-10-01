/* Actual native subsampling, frame-edge padding, mean subtraction and CfL prediction. */
void office_cfl_prediction(const int *p, const uint8_t *luma, uint8_t *output) {
  CFL_CTX cfl; memset(&cfl, 0, sizeof(cfl));
  if (p[3] && p[4]) cfl_luma_subsampling_420_lbd_c(luma, p[7], cfl.recon_buf_q3, p[5], p[6]);
  else if (p[3]) cfl_luma_subsampling_422_lbd_c(luma, p[7], cfl.recon_buf_q3, p[5], p[6]);
  else cfl_luma_subsampling_444_lbd_c(luma, p[7], cfl.recon_buf_q3, p[5], p[6]);
  cfl.buf_width = p[5] >> p[3]; cfl.buf_height = p[6] >> p[4];
  int w = tx_size_wide[p[0]], h = tx_size_high[p[0]];
  cfl_pad(&cfl, w, h);
  subtract_average_c(cfl.recon_buf_q3, cfl.ac_buf_q3, w, h, w * h / 2,
      tx_size_wide_log2[p[0]] + tx_size_high_log2[p[0]]);
  memset(output, p[2], (size_t)w * h);
  cfl_predict_lbd_c(cfl.ac_buf_q3, output, w, p[1], w, h);
}
