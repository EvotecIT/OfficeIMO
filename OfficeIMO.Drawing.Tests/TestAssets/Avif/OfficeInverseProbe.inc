/* Export the real native inverse routines to the opt-in fixture driver. */
void office_av1_inverse(const int32_t *input,int32_t *residual,uint16_t *pixels,int size,int type,int lossless,int bit_depth) {
  typedef void (*inverse)(const int32_t *,uint16_t *,int,TX_TYPE,int);
  static const inverse functions[]={
    av1_inv_txfm2d_add_4x4_c,av1_inv_txfm2d_add_8x8_c,av1_inv_txfm2d_add_16x16_c,av1_inv_txfm2d_add_32x32_c,av1_inv_txfm2d_add_64x64_c,
    av1_inv_txfm2d_add_4x8_c,av1_inv_txfm2d_add_8x4_c,av1_inv_txfm2d_add_8x16_c,av1_inv_txfm2d_add_16x8_c,av1_inv_txfm2d_add_16x32_c,
    av1_inv_txfm2d_add_32x16_c,av1_inv_txfm2d_add_32x64_c,av1_inv_txfm2d_add_64x32_c,av1_inv_txfm2d_add_4x16_c,av1_inv_txfm2d_add_16x4_c,
    av1_inv_txfm2d_add_8x32_c,av1_inv_txfm2d_add_32x8_c,av1_inv_txfm2d_add_16x64_c,av1_inv_txfm2d_add_64x16_c
  };
  extern int office_probe_range_valid;office_probe_range_valid=1;
  office_probe_residual=residual;office_probe_origin=pixels;
  if(lossless) av1_highbd_iwht4x4_16_add_c(input,CONVERT_TO_BYTEPTR(pixels),4,bit_depth);
  else functions[size](input,pixels,tx_size_wide[size],(TX_TYPE)type,bit_depth);
  office_probe_residual=NULL;office_probe_origin=NULL;
}
