import bcw2sv
bcw2sv.extractReg()
bcw2sv.genUpdateBcur()
bcw2sv.genBcwName2Addr()
bcw2sv.genReadBcwDefault()
bcw2sv.genBcwWriteRead()
bcw2sv.genBcurBcw()
bcw2sv.genBcwWw()
import comb_sv
comb_sv.run()
import os
os.system('gvim db_auto.sv')
