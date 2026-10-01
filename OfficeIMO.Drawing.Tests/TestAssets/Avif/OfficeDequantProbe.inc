/* The fixture uses the native dequant factor/matrix lookup and the decoder's Main-8 coefficient storage rule. */
int office_av1_dequant(int level,int dc,int ac,int position,int weight,int size) {
  const int16_t dequant[]={(int16_t)dc,(int16_t)ac};
  const qm_val_t matrix[]={(qm_val_t)weight,(qm_val_t)weight};
  int factor=get_dqv(dequant,position?1:0,weight?matrix:NULL);
  int64_t magnitude=level<0?-(int64_t)level:level;
  int value=(int)((magnitude*factor & 0xffffff)>>av1_get_tx_scale((TX_SIZE)size));
  return clamp(level<0?-value:value,-32768,32767);
}
