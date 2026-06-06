"""
========================================================================
  GUJARAT STATE LOAD DESPATCH CENTRE
  Security Constrained Unit Commitment (SCUC) +
  Security Constrained Economic Dispatch (SCED)
  GAMSPy Implementation - Full 96-block MIP
  Date: 2026-06-05
========================================================================

REQUIREMENTS:
  pip install gamspy pandas openpyxl

USAGE:
  python Gujarat_SCUC_SCED_GAMSPy.py

OUTPUT:
  Gujarat_SCUC_Results.xlsx  - Full dispatch/commitment/cost tables
  Gujarat_SCUC_Results.csv   - Flat results for further analysis
========================================================================
"""

# ======================================================================
#  SECTION 1 - IMPORTS
# ======================================================================

from gamspy import (
    Container, Set, Alias, Parameter, Variable, Equation, Model, Sum,
    Smax, Smin, Ord, Card, Number
)
from gamspy import Problem, Sense
import pandas as pd
import numpy as np

# ======================================================================
#  SECTION 2 - RAW DATA (sourced from Gujarat_SCUC_SCED_Phase1)
# ======================================================================

GENERATORS = [
    "DHUVARAN_CCPP", "UKAI_TPS", "GANDHINAGAR_TPS", "WANAKBORI_TPS",
    "SIKKA_TPS", "PANANDHRO", "BECL", "UTRAN_II_CCPP", "UKAI_HYDRO",
    "KADANA_HYDRO", "TP_AECO_AMGEN", "GIPCL_I_MANGROL", "SLPP_GIPCL",
    "GSEG_II_GSPC", "EPGL_ESSAR", "AKRIMOTA_BECL", "SUGEN_CCPP",
    "UNOSUGEN", "APL_ADANI_MUNDRA", "ACB_INDIA_LTD", "GSPC_PIPAVAV_CLPI",
    "DB_POWER", "KAKRAPAR_KAPS_12", "KAKRAPAR_EXT_34", "TARAPUR_12",
    "TARAPUR_34", "KORBA_NTPC", "KORBA_7_NTPC", "VSTPS_I", "VSTPS_II",
    "VSTPS_III", "VSTPS_IV", "VSTPS_V", "JHANOR_JGPS_KAWAS", "KAWAS_GAS",
    "SIPAT_I", "SIPAT_II", "GADARWARA_I", "CGPL_MUNDRA_UMPP",
    "MSTPS_I_MAUDA", "MSTPS_II_MAUDA", "LARA_STPP", "KHARGONE_STPS",
    "DADRI_I", "SSP_CHPH", "SSP_RBPH", "BIOMASS",
]
ZONES   = ["KUTCH", "SAURASHTRA", "N_GUJ", "CENTRAL", "S_GUJ"]
BLOCKS  = [f"T{str(k).zfill(2)}" for k in range(1, 97)]
MUST_RUN = ["KAKRAPAR_KAPS_12", "KAKRAPAR_EXT_34", "TARAPUR_12",
            "TARAPUR_34", "SSP_CHPH", "SSP_RBPH"]

# --- Generator parameters [AUTH/PROXY: Phase 1 Sheets 1-6, 24] ---
PMAX = {
    "DHUVARAN_CCPP":488.1,"UKAI_TPS":1110.0,"GANDHINAGAR_TPS":630.0,
    "WANAKBORI_TPS":2270.0,"SIKKA_TPS":500.0,"PANANDHRO":150.0,
    "BECL":500.0,"UTRAN_II_CCPP":374.57,"UKAI_HYDRO":305.0,
    "KADANA_HYDRO":242.0,"TP_AECO_AMGEN":362.0,"GIPCL_I_MANGROL":93.0,
    "SLPP_GIPCL":500.0,"GSEG_II_GSPC":351.0,"EPGL_ESSAR":1122.0,
    "AKRIMOTA_BECL":250.0,"SUGEN_CCPP":1148.0,"UNOSUGEN":383.0,
    "APL_ADANI_MUNDRA":2434.0,"ACB_INDIA_LTD":200.0,
    "GSPC_PIPAVAV_CLPI":703.0,"DB_POWER":293.0,"KAKRAPAR_KAPS_12":125.0,
    "KAKRAPAR_EXT_34":475.88,"TARAPUR_12":160.0,"TARAPUR_34":274.0,
    "KORBA_NTPC":380.84,"KORBA_7_NTPC":130.34,"VSTPS_I":247.74,
    "VSTPS_II":252.64,"VSTPS_III":279.64,"VSTPS_IV":258.6,"VSTPS_V":102.79,
    "JHANOR_JGPS_KAWAS":237.0,"KAWAS_GAS":187.0,"SIPAT_I":576.99,
    "SIPAT_II":286.02,"GADARWARA_I":333.83,"CGPL_MUNDRA_UMPP":1805.0,
    "MSTPS_I_MAUDA":455.633,"MSTPS_II_MAUDA":530.243,"LARA_STPP":186.478,
    "KHARGONE_STPS":270.197,"DADRI_I":576.996,"SSP_CHPH":40.0,
    "SSP_RBPH":192.0,"BIOMASS":81.65,
}
PMIN = {
    "DHUVARAN_CCPP":30.0,"UKAI_TPS":880.6,"GANDHINAGAR_TPS":436.0,
    "WANAKBORI_TPS":1013.6,"SIKKA_TPS":330.0,"PANANDHRO":45.0,
    "BECL":150.0,"UTRAN_II_CCPP":112.371,"UKAI_HYDRO":91.5,
    "KADANA_HYDRO":72.6,"TP_AECO_AMGEN":108.6,"GIPCL_I_MANGROL":27.9,
    "SLPP_GIPCL":240.0,"GSEG_II_GSPC":105.3,"EPGL_ESSAR":720.0,
    "AKRIMOTA_BECL":100.0,"SUGEN_CCPP":30.0,"UNOSUGEN":30.0,
    "APL_ADANI_MUNDRA":730.2,"ACB_INDIA_LTD":60.0,
    "GSPC_PIPAVAV_CLPI":210.9,"DB_POWER":87.9,"KAKRAPAR_KAPS_12":37.5,
    "KAKRAPAR_EXT_34":142.764,"TARAPUR_12":48.0,"TARAPUR_34":82.2,
    "KORBA_NTPC":114.252,"KORBA_7_NTPC":39.102,"VSTPS_I":74.322,
    "VSTPS_II":75.792,"VSTPS_III":83.892,"VSTPS_IV":77.58,"VSTPS_V":30.837,
    "JHANOR_JGPS_KAWAS":71.1,"KAWAS_GAS":56.1,"SIPAT_I":173.097,
    "SIPAT_II":85.806,"GADARWARA_I":100.149,"CGPL_MUNDRA_UMPP":541.5,
    "MSTPS_I_MAUDA":136.69,"MSTPS_II_MAUDA":159.073,"LARA_STPP":55.943,
    "KHARGONE_STPS":81.059,"DADRI_I":173.099,"SSP_CHPH":12.0,
    "SSP_RBPH":57.6,"BIOMASS":24.495,
}
VC = {
    "DHUVARAN_CCPP":13.08,"UKAI_TPS":5.35,"GANDHINAGAR_TPS":5.39,
    "WANAKBORI_TPS":5.23,"SIKKA_TPS":6.1,"PANANDHRO":4.23,"BECL":4.23,
    "UTRAN_II_CCPP":12.73,"UKAI_HYDRO":2.05,"KADANA_HYDRO":2.05,
    "TP_AECO_AMGEN":4.23,"GIPCL_I_MANGROL":3.36,"SLPP_GIPCL":3.36,
    "GSEG_II_GSPC":10.3,"EPGL_ESSAR":5.52,"AKRIMOTA_BECL":2.84,
    "SUGEN_CCPP":4.94,"UNOSUGEN":3.79,"APL_ADANI_MUNDRA":5.16,
    "ACB_INDIA_LTD":1.09,"GSPC_PIPAVAV_CLPI":4.94,"DB_POWER":1.56,
    "KAKRAPAR_KAPS_12":2.29,"KAKRAPAR_EXT_34":4.4,"TARAPUR_12":2.42,
    "TARAPUR_34":3.65,"KORBA_NTPC":1.39,"KORBA_7_NTPC":1.37,
    "VSTPS_I":2.23,"VSTPS_II":2.14,"VSTPS_III":2.12,"VSTPS_IV":2.1,
    "VSTPS_V":2.16,"JHANOR_JGPS_KAWAS":7.25,"KAWAS_GAS":7.28,
    "SIPAT_I":1.27,"SIPAT_II":1.31,"GADARWARA_I":3.24,
    "CGPL_MUNDRA_UMPP":4.6,"MSTPS_I_MAUDA":3.19,"MSTPS_II_MAUDA":3.15,
    "LARA_STPP":1.31,"KHARGONE_STPS":3.41,"DADRI_I":4.26,
    "SSP_CHPH":2.05,"SSP_RBPH":2.05,"BIOMASS":5.0,
}
RAMPUP = {
    "DHUVARAN_CCPP":282.0,"UKAI_TPS":687.0,"GANDHINAGAR_TPS":37.5,
    "WANAKBORI_TPS":37.5,"SIKKA_TPS":37.5,"PANANDHRO":19.5,"BECL":19.5,
    "UTRAN_II_CCPP":282.0,"UKAI_HYDRO":544.5,"KADANA_HYDRO":544.5,
    "TP_AECO_AMGEN":37.5,"GIPCL_I_MANGROL":687.0,"SLPP_GIPCL":19.5,
    "GSEG_II_GSPC":282.0,"EPGL_ESSAR":180.0,"AKRIMOTA_BECL":19.5,
    "SUGEN_CCPP":282.0,"UNOSUGEN":861.0,"APL_ADANI_MUNDRA":297.0,
    "ACB_INDIA_LTD":31.5,"GSPC_PIPAVAV_CLPI":300.0,"DB_POWER":31.5,
    "KAKRAPAR_KAPS_12":7.5,"KAKRAPAR_EXT_34":7.5,"TARAPUR_12":7.5,
    "TARAPUR_34":7.5,"KORBA_NTPC":31.5,"KORBA_7_NTPC":31.5,
    "VSTPS_I":687.0,"VSTPS_II":282.0,"VSTPS_III":31.5,"VSTPS_IV":31.5,
    "VSTPS_V":3262.5,"JHANOR_JGPS_KAWAS":789.0,"KAWAS_GAS":789.0,
    "SIPAT_I":687.0,"SIPAT_II":282.0,"GADARWARA_I":687.0,
    "CGPL_MUNDRA_UMPP":373.5,"MSTPS_I_MAUDA":687.0,"MSTPS_II_MAUDA":282.0,
    "LARA_STPP":31.5,"KHARGONE_STPS":31.5,"DADRI_I":687.0,
    "SSP_CHPH":3262.5,"SSP_RBPH":3262.5,"BIOMASS":31.5,
}
RAMPDN = RAMPUP.copy()  # symmetric ramp rates

SU_COST = {
    "DHUVARAN_CCPP":4959622,"UKAI_TPS":30511403,"GANDHINAGAR_TPS":14614700,
    "WANAKBORI_TPS":33062908,"SIKKA_TPS":12614250,"PANANDHRO":75000,
    "BECL":250000,"UTRAN_II_CCPP":187285,"UKAI_HYDRO":152500,
    "KADANA_HYDRO":121000,"TP_AECO_AMGEN":181000,"GIPCL_I_MANGROL":46500,
    "SLPP_GIPCL":5293200,"GSEG_II_GSPC":175500,"EPGL_ESSAR":32965200,
    "AKRIMOTA_BECL":1866500,"SUGEN_CCPP":3962322,"UNOSUGEN":1057654,
    "APL_ADANI_MUNDRA":1217000,"ACB_INDIA_LTD":100000,
    "GSPC_PIPAVAV_CLPI":351500,"DB_POWER":146500,"KAKRAPAR_KAPS_12":62500,
    "KAKRAPAR_EXT_34":237940,"TARAPUR_12":80000,"TARAPUR_34":137000,
    "KORBA_NTPC":190420,"KORBA_7_NTPC":65170,"VSTPS_I":123870,
    "VSTPS_II":126320,"VSTPS_III":139820,"VSTPS_IV":129300,"VSTPS_V":51395,
    "JHANOR_JGPS_KAWAS":118500,"KAWAS_GAS":93500,"SIPAT_I":288495,
    "SIPAT_II":143010,"GADARWARA_I":166915,"CGPL_MUNDRA_UMPP":902500,
    "MSTPS_I_MAUDA":227816,"MSTPS_II_MAUDA":265122,"LARA_STPP":13986,
    "KHARGONE_STPS":20265,"DADRI_I":43275,"SSP_CHPH":3000,
    "SSP_RBPH":14400,"BIOMASS":6124,  # using NoLoad values as proxy for missing
}
NOLOAD = {
    "DHUVARAN_CCPP":74394,"UKAI_TPS":457671,"GANDHINAGAR_TPS":219221,
    "WANAKBORI_TPS":495944,"SIKKA_TPS":189214,"PANANDHRO":11250,
    "BECL":37500,"UTRAN_II_CCPP":28093,"UKAI_HYDRO":22875,
    "KADANA_HYDRO":18150,"TP_AECO_AMGEN":27150,"GIPCL_I_MANGROL":6975,
    "SLPP_GIPCL":793980,"GSEG_II_GSPC":26325,"EPGL_ESSAR":4944780,
    "AKRIMOTA_BECL":279975,"SUGEN_CCPP":594348,"UNOSUGEN":158648,
    "APL_ADANI_MUNDRA":182550,"ACB_INDIA_LTD":15000,
    "GSPC_PIPAVAV_CLPI":52725,"DB_POWER":21975,"KAKRAPAR_KAPS_12":9375,
    "KAKRAPAR_EXT_34":35691,"TARAPUR_12":12000,"TARAPUR_34":20550,
    "KORBA_NTPC":28563,"KORBA_7_NTPC":9776,"VSTPS_I":18581,
    "VSTPS_II":18948,"VSTPS_III":20973,"VSTPS_IV":19395,"VSTPS_V":7709,
    "JHANOR_JGPS_KAWAS":17775,"KAWAS_GAS":14025,"SIPAT_I":43274,
    "SIPAT_II":21452,"GADARWARA_I":25037,"CGPL_MUNDRA_UMPP":135375,
    "MSTPS_I_MAUDA":34172,"MSTPS_II_MAUDA":39768,"LARA_STPP":13986,
    "KHARGONE_STPS":20265,"DADRI_I":43275,"SSP_CHPH":3000,
    "SSP_RBPH":14400,"BIOMASS":6124,
}
MUT = {
    "DHUVARAN_CCPP":8,"UKAI_TPS":32,"GANDHINAGAR_TPS":24,"WANAKBORI_TPS":24,
    "SIKKA_TPS":24,"PANANDHRO":8,"BECL":24,"UTRAN_II_CCPP":8,
    "UKAI_HYDRO":8,"KADANA_HYDRO":8,"TP_AECO_AMGEN":24,"GIPCL_I_MANGROL":8,
    "SLPP_GIPCL":24,"GSEG_II_GSPC":8,"EPGL_ESSAR":32,"AKRIMOTA_BECL":24,
    "SUGEN_CCPP":8,"UNOSUGEN":8,"APL_ADANI_MUNDRA":24,"ACB_INDIA_LTD":24,
    "GSPC_PIPAVAV_CLPI":8,"DB_POWER":24,"KAKRAPAR_KAPS_12":8,
    "KAKRAPAR_EXT_34":8,"TARAPUR_12":8,"TARAPUR_34":8,"KORBA_NTPC":24,
    "KORBA_7_NTPC":24,"VSTPS_I":24,"VSTPS_II":24,"VSTPS_III":24,
    "VSTPS_IV":24,"VSTPS_V":24,"JHANOR_JGPS_KAWAS":8,"KAWAS_GAS":8,
    "SIPAT_I":24,"SIPAT_II":24,"GADARWARA_I":24,"CGPL_MUNDRA_UMPP":24,
    "MSTPS_I_MAUDA":24,"MSTPS_II_MAUDA":24,"LARA_STPP":24,
    "KHARGONE_STPS":24,"DADRI_I":24,"SSP_CHPH":8,"SSP_RBPH":8,"BIOMASS":24,
}
MDT = {
    "DHUVARAN_CCPP":8,"UKAI_TPS":24,"GANDHINAGAR_TPS":16,"WANAKBORI_TPS":16,
    "SIKKA_TPS":16,"PANANDHRO":8,"BECL":16,"UTRAN_II_CCPP":8,
    "UKAI_HYDRO":8,"KADANA_HYDRO":8,"TP_AECO_AMGEN":16,"GIPCL_I_MANGROL":8,
    "SLPP_GIPCL":16,"GSEG_II_GSPC":8,"EPGL_ESSAR":24,"AKRIMOTA_BECL":16,
    "SUGEN_CCPP":8,"UNOSUGEN":8,"APL_ADANI_MUNDRA":16,"ACB_INDIA_LTD":16,
    "GSPC_PIPAVAV_CLPI":8,"DB_POWER":16,"KAKRAPAR_KAPS_12":8,
    "KAKRAPAR_EXT_34":8,"TARAPUR_12":8,"TARAPUR_34":8,"KORBA_NTPC":16,
    "KORBA_7_NTPC":16,"VSTPS_I":16,"VSTPS_II":16,"VSTPS_III":16,
    "VSTPS_IV":16,"VSTPS_V":16,"JHANOR_JGPS_KAWAS":8,"KAWAS_GAS":8,
    "SIPAT_I":16,"SIPAT_II":16,"GADARWARA_I":16,"CGPL_MUNDRA_UMPP":16,
    "MSTPS_I_MAUDA":16,"MSTPS_II_MAUDA":16,"LARA_STPP":16,
    "KHARGONE_STPS":16,"DADRI_I":16,"SSP_CHPH":8,"SSP_RBPH":8,"BIOMASS":16,
}
GEN_ZONE = {
    "DHUVARAN_CCPP":"S_GUJ","UKAI_TPS":"S_GUJ","GANDHINAGAR_TPS":"CENTRAL",
    "WANAKBORI_TPS":"CENTRAL","SIKKA_TPS":"SAURASHTRA","PANANDHRO":"KUTCH",
    "BECL":"KUTCH","UTRAN_II_CCPP":"S_GUJ","UKAI_HYDRO":"S_GUJ",
    "KADANA_HYDRO":"CENTRAL","TP_AECO_AMGEN":"KUTCH","GIPCL_I_MANGROL":"S_GUJ",
    "SLPP_GIPCL":"S_GUJ","GSEG_II_GSPC":"S_GUJ","EPGL_ESSAR":"KUTCH",
    "AKRIMOTA_BECL":"KUTCH","SUGEN_CCPP":"S_GUJ","UNOSUGEN":"S_GUJ",
    "APL_ADANI_MUNDRA":"KUTCH","ACB_INDIA_LTD":"KUTCH",
    "GSPC_PIPAVAV_CLPI":"SAURASHTRA","DB_POWER":"CENTRAL",
    "KAKRAPAR_KAPS_12":"S_GUJ","KAKRAPAR_EXT_34":"S_GUJ",
    "TARAPUR_12":"S_GUJ","TARAPUR_34":"S_GUJ","KORBA_NTPC":"CENTRAL",
    "KORBA_7_NTPC":"CENTRAL","VSTPS_I":"CENTRAL","VSTPS_II":"CENTRAL",
    "VSTPS_III":"CENTRAL","VSTPS_IV":"CENTRAL","VSTPS_V":"CENTRAL",
    "JHANOR_JGPS_KAWAS":"S_GUJ","KAWAS_GAS":"S_GUJ","SIPAT_I":"CENTRAL",
    "SIPAT_II":"CENTRAL","GADARWARA_I":"CENTRAL","CGPL_MUNDRA_UMPP":"KUTCH",
    "MSTPS_I_MAUDA":"CENTRAL","MSTPS_II_MAUDA":"CENTRAL",
    "LARA_STPP":"CENTRAL","KHARGONE_STPS":"CENTRAL","DADRI_I":"CENTRAL",
    "SSP_CHPH":"CENTRAL","SSP_RBPH":"CENTRAL","BIOMASS":"CENTRAL",
}
TRANSFER = {
    ("KUTCH","SAURASHTRA"):1000, ("KUTCH","CENTRAL"):2800,   # +300 MW: Derol-Shamlaji 400kV
    ("SAURASHTRA","KUTCH"):1000, ("SAURASHTRA","CENTRAL"):1250,
    ("N_GUJ","CENTRAL"):3500,
    ("CENTRAL","KUTCH"):2800,                                 # +300 MW bidirectional
    ("CENTRAL","SAURASHTRA"):1250,
    ("CENTRAL","N_GUJ"):3500,
    ("CENTRAL","S_GUJ"):2800,                                 # +300 MW: Asoj-Vadodara 400kV
    ("S_GUJ","CENTRAL"):2800,                                 # +300 MW bidirectional
    ("N_GUJ","KUTCH"):500,
    ("KUTCH","N_GUJ"):500,
    ("N_GUJ","SAURASHTRA"):500,
    ("SAURASHTRA","N_GUJ"):500,
    ("S_GUJ","SAURASHTRA"):500,                               # added: S_GUJ-Saurashtra link
    ("SAURASHTRA","S_GUJ"):500,
}

# --- 96-block demand profile [REPR: Sheet21] ---
DEMAND_T = {
    "T01":20091,"T02":20050,"T03":20000,"T04":19940,"T05":19871,"T06":19794,
    "T07":19714,"T08":19634,"T09":19559,"T10":19494,"T11":19443,"T12":19411,
    "T13":19400,"T14":19411,"T15":19444,"T16":19495,"T17":19562,"T18":19640,
    "T19":19725,"T20":19813,"T21":19902,"T22":19990,"T23":20080,"T24":20174,
    "T25":20275,"T26":20391,"T27":20526,"T28":20686,"T29":20875,"T30":21094,
    "T31":21342,"T32":21612,"T33":21895,"T34":22176,"T35":22440,"T36":22669,
    "T37":22845,"T38":22955,"T39":22986,"T40":22933,"T41":22795,"T42":22575,
    "T43":22279,"T44":21919,"T45":21508,"T46":21068,"T47":20623,"T48":20203,
    "T49":19839,"T50":19558,"T51":19378,"T52":19304,"T53":19325,"T54":19417,
    "T55":19552,"T56":19701,"T57":19841,"T58":19959,"T59":20048,"T60":20110,
    "T61":20151,"T62":20176,"T63":20191,"T64":20202,"T65":20213,"T66":20230,
    "T67":20259,"T68":20307,"T69":20384,"T70":20503,"T71":20679,"T72":20924,
    "T73":21247,"T74":21650,"T75":22122,"T76":22641,"T77":23167,"T78":23654,
    "T79":24050,"T80":24309,"T81":24400,"T82":24309,"T83":24050,"T84":23654,
    "T85":23167,"T86":22641,"T87":22122,"T88":21650,"T89":21247,"T90":20924,
    "T91":20679,"T92":20503,"T93":20384,"T94":20307,"T95":20259,"T96":20231,
}

# --- RE profiles [REPR: Sheet22] ---
SOLAR_KUTCH = {t: 0 for t in BLOCKS}
_sk = {"T25":163,"T26":198,"T27":238,"T28":283,"T29":335,"T30":393,"T31":458,
       "T32":529,"T33":606,"T34":688,"T35":776,"T36":867,"T37":962,"T38":1059,
       "T39":1157,"T40":1253,"T41":1346,"T42":1435,"T43":1518,"T44":1593,
       "T45":1657,"T46":1711,"T47":1753,"T48":1782,"T49":1796,"T50":1796,
       "T51":1783,"T52":1755,"T53":1713,"T54":1660,"T55":1595,"T56":1521,
       "T57":1439,"T58":1350,"T59":1257,"T60":1161,"T61":1063,"T62":966,
       "T63":871,"T64":779,"T65":691,"T66":609,"T67":532,"T68":460,"T69":396,
       "T70":337,"T71":285,"T72":239,"T73":199,"T74":164,"T75":135}
SOLAR_KUTCH.update(_sk)

SOLAR_NGUJ = {t: 0 for t in BLOCKS}
_sn = {"T25":318,"T26":385,"T27":463,"T28":551,"T29":652,"T30":765,"T31":891,
       "T32":1028,"T33":1178,"T34":1338,"T35":1509,"T36":1687,"T37":1872,
       "T38":2060,"T39":2249,"T40":2437,"T41":2618,"T42":2791,"T43":2952,
       "T44":3097,"T45":3223,"T46":3328,"T47":3410,"T48":3465,"T49":3493,
       "T50":3494,"T51":3467,"T52":3412,"T53":3332,"T54":3228,"T55":3102,
       "T56":2958,"T57":2798,"T58":2625,"T59":2444,"T60":2257,"T61":2068,
       "T62":1879,"T63":1695,"T64":1516,"T65":1345,"T66":1184,"T67":1034,
       "T68":896,"T69":770,"T70":656,"T71":555,"T72":466,"T73":388,
       "T74":320,"T75":262}
SOLAR_NGUJ.update(_sn)

SOLAR_CENT = {t: 0 for t in BLOCKS}
_sc = {"T25":227,"T26":275,"T27":330,"T28":394,"T29":466,"T30":547,"T31":636,
       "T32":734,"T33":841,"T34":956,"T35":1078,"T36":1205,"T37":1337,
       "T38":1471,"T39":1607,"T40":1740,"T41":1870,"T42":1994,"T43":2108,
       "T44":2212,"T45":2302,"T46":2377,"T47":2435,"T48":2475,"T49":2495,
       "T50":2495,"T51":2476,"T52":2437,"T53":2380,"T54":2306,"T55":2216,
       "T56":2113,"T57":1998,"T58":1875,"T59":1746,"T60":1612,"T61":1477,
       "T62":1342,"T63":1210,"T64":1083,"T65":961,"T66":846,"T67":739,
       "T68":640,"T69":550,"T70":469,"T71":397,"T72":333,"T73":277,
       "T74":229,"T75":187}
SOLAR_CENT.update(_sc)

SOLAR_SGUJ = {t: 0 for t in BLOCKS}
_ss = {"T25":201,"T26":243,"T27":292,"T28":349,"T29":412,"T30":483,"T31":561,
       "T32":648,"T33":742,"T34":843,"T35":949,"T36":1063,"T37":1178,
       "T38":1297,"T39":1415,"T40":1533,"T41":1648,"T42":1756,"T43":1857,
       "T44":1948,"T45":2029,"T46":2095,"T47":2145,"T48":2179,"T49":2198,
       "T50":2198,"T51":2180,"T52":2146,"T53":2097,"T54":2030,"T55":1952,
       "T56":1860,"T57":1760,"T58":1652,"T59":1537,"T60":1420,"T61":1301,
       "T62":1184,"T63":1067,"T64":954,"T65":847,"T66":745,"T67":651,
       "T68":565,"T69":485,"T70":415,"T71":351,"T72":295,"T73":246,
       "T74":203,"T75":167}
SOLAR_SGUJ.update(_ss)

SOLAR_SAUR = {t: 0 for t in BLOCKS}

WIND_KUTCH = {
    "T01":2249,"T02":2249,"T03":2249,"T04":2249,"T05":2249,"T06":2249,
    "T07":2249,"T08":2249,"T09":2248,"T10":2248,"T11":2248,"T12":2247,
    "T13":2246,"T14":2245,"T15":2244,"T16":2242,"T17":2241,"T18":2238,
    "T19":2235,"T20":2231,"T21":2226,"T22":2221,"T23":2214,"T24":2205,
    "T25":2196,"T26":2185,"T27":2172,"T28":2157,"T29":2140,"T30":2120,
    "T31":2098,"T32":2074,"T33":2047,"T34":2018,"T35":1986,"T36":1953,
    "T37":1916,"T38":1878,"T39":1839,"T40":1799,"T41":1758,"T42":1717,
    "T43":1677,"T44":1638,"T45":1601,"T46":1566,"T47":1534,"T48":1507,
    "T49":1483,"T50":1464,"T51":1451,"T52":1442,"T53":1440,"T54":1443,
    "T55":1452,"T56":1466,"T57":1486,"T58":1511,"T59":1540,"T60":1575,
    "T61":1613,"T62":1655,"T63":1701,"T64":1749,"T65":1801,"T66":1855,
    "T67":1912,"T68":1971,"T69":2033,"T70":2097,"T71":2162,"T72":2228,
    "T73":2295,"T74":2361,"T75":2426,"T76":2488,"T77":2547,"T78":2601,
    "T79":2648,"T80":2688,"T81":2719,"T82":2741,"T83":2754,"T84":2756,
    "T85":2749,"T86":2734,"T87":2711,"T88":2682,"T89":2648,"T90":2610,
    "T91":2571,"T92":2532,"T93":2493,"T94":2457,"T95":2423,"T96":2392,
}
WIND_SAUR = {
    "T01":1999,"T02":1999,"T03":1999,"T04":1999,"T05":1999,"T06":1999,
    "T07":1999,"T08":1999,"T09":1998,"T10":1998,"T11":1998,"T12":1997,
    "T13":1997,"T14":1996,"T15":1995,"T16":1993,"T17":1992,"T18":1989,
    "T19":1986,"T20":1983,"T21":1979,"T22":1974,"T23":1968,"T24":1960,
    "T25":1952,"T26":1942,"T27":1930,"T28":1917,"T29":1902,"T30":1885,
    "T31":1865,"T32":1844,"T33":1820,"T34":1794,"T35":1766,"T36":1736,
    "T37":1703,"T38":1670,"T39":1635,"T40":1599,"T41":1563,"T42":1526,
    "T43":1491,"T44":1456,"T45":1423,"T46":1392,"T47":1364,"T48":1339,
    "T49":1318,"T50":1302,"T51":1290,"T52":1282,"T53":1280,"T54":1282,
    "T55":1290,"T56":1303,"T57":1321,"T58":1343,"T59":1369,"T60":1400,
    "T61":1434,"T62":1471,"T63":1512,"T64":1555,"T65":1601,"T66":1649,
    "T67":1700,"T68":1752,"T69":1807,"T70":1864,"T71":1922,"T72":1980,
    "T73":2040,"T74":2099,"T75":2156,"T76":2212,"T77":2264,"T78":2312,
    "T79":2354,"T80":2390,"T81":2417,"T82":2437,"T83":2448,"T84":2450,
    "T85":2444,"T86":2430,"T87":2410,"T88":2384,"T89":2354,"T90":2320,
    "T91":2286,"T92":2251,"T93":2216,"T94":2184,"T95":2154,"T96":2126,
}

# Derived: wind in N_GUJ, CENTRAL, S_GUJ from Wind_Total
WIND_TOTAL = {
    "T01":4248,"T02":4248,"T03":4248,"T04":4248,"T05":4248,"T06":4248,
    "T07":4248,"T08":4248,"T09":4246,"T10":4246,"T11":4246,"T12":4244,
    "T13":4243,"T14":4241,"T15":4239,"T16":4235,"T17":4233,"T18":4227,
    "T19":4221,"T20":4214,"T21":4205,"T22":4195,"T23":4182,"T24":4165,
    "T25":4148,"T26":4127,"T27":4102,"T28":4074,"T29":4042,"T30":4005,
    "T31":3963,"T32":3918,"T33":3867,"T34":3812,"T35":3752,"T36":3689,
    "T37":3619,"T38":3548,"T39":3474,"T40":3398,"T41":3321,"T42":3243,
    "T43":3168,"T44":3094,"T45":3024,"T46":2958,"T47":2898,"T48":2846,
    "T49":2801,"T50":2766,"T51":2741,"T52":2724,"T53":2720,"T54":2725,
    "T55":2742,"T56":2769,"T57":2807,"T58":2854,"T59":2909,"T60":2975,
    "T61":3047,"T62":3126,"T63":3213,"T64":3304,"T65":3402,"T66":3504,
    "T67":3612,"T68":3723,"T69":3840,"T70":3961,"T71":4084,"T72":4208,
    "T73":4335,"T74":4460,"T75":4582,"T76":4700,"T77":4811,"T78":4913,
    "T79":5002,"T80":5078,"T81":5136,"T82":5178,"T83":5202,"T84":5206,
    "T85":5193,"T86":5164,"T87":5121,"T88":5066,"T89":5002,"T90":4930,
    "T91":4857,"T92":4783,"T93":4709,"T94":4641,"T95":4577,"T96":4518,
}
WIND_NGUJ    = {t: 0.0750 * WIND_TOTAL[t] / 0.85 for t in BLOCKS}
WIND_CENTRAL = {t: 0.0375 * WIND_TOTAL[t] / 0.85 for t in BLOCKS}
WIND_SGUJ    = {t: 0.0375 * WIND_TOTAL[t] / 0.85 for t in BLOCKS}

# RE available per zone per block
RE_AVAIL = {
    ("KUTCH",    t): SOLAR_KUTCH[t] + WIND_KUTCH[t]   for t in BLOCKS
}
RE_AVAIL.update({("SAURASHTRA", t): SOLAR_SAUR[t]  + WIND_SAUR[t]    for t in BLOCKS})
RE_AVAIL.update({("N_GUJ",      t): SOLAR_NGUJ[t]  + WIND_NGUJ[t]    for t in BLOCKS})
RE_AVAIL.update({("CENTRAL",    t): SOLAR_CENT[t]  + WIND_CENTRAL[t] for t in BLOCKS})
RE_AVAIL.update({("S_GUJ",      t): SOLAR_SGUJ[t]  + WIND_SGUJ[t]    for t in BLOCKS})

# Unserved energy penalty (Value of Lost Load = Rs 50,000/MWh)
# This ensures model is always feasible - gap shows up as cost not infeasibility
VOLL = 50000  # Rs/MWh = Rs 50/kWh (standard CERC VOLL)

# Zonal demand shares [REPR: Sheet15]
ZONE_SHARE = {"KUTCH":0.12,"SAURASHTRA":0.18,"N_GUJ":0.14,"CENTRAL":0.35,"S_GUJ":0.21}
DEMAND_ZT  = {(z, t): ZONE_SHARE[z] * DEMAND_T[t] for z in ZONES for t in BLOCKS}

# Spinning reserve: 5% of total demand
SPIN_RES_PC = 0.05
SPIN_REQ    = {t: SPIN_RES_PC * DEMAND_T[t] for t in BLOCKS}

# BigM curtailment penalty
BIG_M = 10000.0

# ======================================================================
#  SECTION 3 - BUILD GAMSPY MODEL
# ======================================================================

print("Building GAMSPy model...")

# Use the real GAMS licence (not the GAMSPy demo licence)
# GAMSPy picks up the licence via the GAMS system directory
import subprocess, sys

# Point GAMSPy to the licensed GAMS installation
m = Container(
    system_directory=r"C:\GAMS\53",
)

# --- Sets ---
i  = Set(m, name="i",  records=GENERATORS,  description="Generators")
t  = Set(m, name="t",  records=BLOCKS,       description="Time blocks")
z  = Set(m, name="z",  records=ZONES,        description="Zones")
zz = Alias(m, name="zz", alias_with=z)
tt = Alias(m, name="tt", alias_with=t)

i_mr = Set(m, name="i_mr", domain=i, records=MUST_RUN, description="Must-run units")

# --- Parameters ---
pmax_p  = Parameter(m, "pmax_p",  domain=i, records=pd.DataFrame([(k,v) for k,v in PMAX.items()],    columns=["i","val"]))
pmin_p  = Parameter(m, "pmin_p",  domain=i, records=pd.DataFrame([(k,v) for k,v in PMIN.items()],    columns=["i","val"]))
vc_p    = Parameter(m, "vc_p",    domain=i, records=pd.DataFrame([(k,v) for k,v in VC.items()],      columns=["i","val"]))
ru_p    = Parameter(m, "ru_p",    domain=i, records=pd.DataFrame([(k,v) for k,v in RAMPUP.items()],  columns=["i","val"]))
rd_p    = Parameter(m, "rd_p",    domain=i, records=pd.DataFrame([(k,v) for k,v in RAMPDN.items()],  columns=["i","val"]))
suc_p   = Parameter(m, "suc_p",   domain=i, records=pd.DataFrame([(k,v) for k,v in SU_COST.items()], columns=["i","val"]))
nl_p    = Parameter(m, "nl_p",    domain=i, records=pd.DataFrame([(k,v) for k,v in NOLOAD.items()],  columns=["i","val"]))
mut_p   = Parameter(m, "mut_p",   domain=i, records=pd.DataFrame([(k,v) for k,v in MUT.items()],     columns=["i","val"]))
mdt_p   = Parameter(m, "mdt_p",   domain=i, records=pd.DataFrame([(k,v) for k,v in MDT.items()],     columns=["i","val"]))

gz_p    = Parameter(m, "gz_p",    domain=[i, z],
                    records=pd.DataFrame([(g, z_) for g, z_ in GEN_ZONE.items()],
                                         columns=["i","z"]).assign(val=1))
tl_p    = Parameter(m, "tl_p",    domain=[z, zz],
                    records=pd.DataFrame([(a, b, v) for (a,b),v in TRANSFER.items()],
                                         columns=["z","zz","val"]))
re_p    = Parameter(m, "re_p",    domain=[z, t],
                    records=pd.DataFrame([(z_, t_, v) for (z_,t_),v in RE_AVAIL.items()],
                                         columns=["z","t","val"]))
dem_p   = Parameter(m, "dem_p",   domain=[z, t],
                    records=pd.DataFrame([(z_, t_, v) for (z_,t_),v in DEMAND_ZT.items()],
                                         columns=["z","t","val"]))
spin_p  = Parameter(m, "spin_p",  domain=t,
                    records=pd.DataFrame([(t_, v) for t_,v in SPIN_REQ.items()],
                                         columns=["t","val"]))

# ======================================================================
#  SECTION 4 - VARIABLES
# ======================================================================

P       = Variable(m, "P",       domain=[i, t], type="positive",     description="Dispatch MW")
Curtail = Variable(m, "Curtail", domain=[z, t], type="positive",     description="RE curtailment MW")
Flow    = Variable(m, "Flow",    domain=[z, zz, t], type="free",     description="Inter-zonal flow MW")
USE     = Variable(m, "USE",     domain=[z, t], type="positive",     description="Unserved energy MW (slack)")
u       = Variable(m, "u",       domain=[i, t], type="binary",       description="Commitment status")
su      = Variable(m, "su",      domain=[i, t], type="binary",       description="Startup event")
sd      = Variable(m, "sd",      domain=[i, t], type="binary",       description="Shutdown event")
TotalCost = Variable(m, "TotalCost", type="free",                    description="Objective Rs")

# Cap unserved energy per zone per block (can't unserve more than demand)
USE.up[z, t] = dem_p[z, t]

# Bounds
P.up[i, t]        = pmax_p[i]
Curtail.up[z, t]  = re_p[z, t]
Flow.lo[z, zz, t] = -tl_p[z, zz]
Flow.up[z, zz, t] =  tl_p[z, zz]

# Fix must-run units
u.fx[i_mr, t]  = 1
su.fx[i_mr, t] = 0
sd.fx[i_mr, t] = 0

# ======================================================================
#  SECTION 5 - EQUATIONS
# ======================================================================

# Objective
eq_obj = Equation(m, "eq_obj", description="Minimise total operating cost")
eq_obj[...] = (
    TotalCost ==
    Sum([i, t], vc_p[i] * 250 * P[i, t])
    + Sum([i, t], nl_p[i] * 0.25 * u[i, t])
    + Sum([i, t], suc_p[i] * su[i, t])
    + Sum([z, t], BIG_M * Curtail[z, t])
    + Sum([z, t], Number(VOLL) * 0.25 * USE[z, t])  # VOLL penalty for unserved energy
)

# Zonal power balance
eq_balance = Equation(m, "eq_balance", domain=[z, t], description="Zonal power balance")
eq_balance[z, t] = (
    Sum(i, gz_p[i, z] * P[i, t])
    + re_p[z, t] - Curtail[z, t]
    + Sum(zz, (tl_p[zz, z] > 0) * Flow[zz, z, t])
    - Sum(zz, (tl_p[z, zz] > 0) * Flow[z, zz, t])
    + USE[z, t]                                      # slack for feasibility
    == dem_p[z, t]
)

# Pmin / Pmax bounds
eq_pmin = Equation(m, "eq_pmin", domain=[i, t], description="Min generation bound")
eq_pmin[i, t] = P[i, t] >= pmin_p[i] * u[i, t]

eq_pmax = Equation(m, "eq_pmax", domain=[i, t], description="Max generation bound")
eq_pmax[i, t] = P[i, t] <= pmax_p[i] * u[i, t]

# Ramp rates (t > T01: Ord(t) > 1)
eq_ramp_up = Equation(m, "eq_ramp_up", domain=[i, t], description="Ramp-up limit")
eq_ramp_up[i, t].where[Ord(t) > 1] = P[i, t] - P[i, t.lag(1)] <= ru_p[i]

eq_ramp_dn = Equation(m, "eq_ramp_dn", domain=[i, t], description="Ramp-down limit")
eq_ramp_dn[i, t].where[Ord(t) > 1] = P[i, t.lag(1)] - P[i, t] <= rd_p[i]

# Startup / shutdown logic
eq_su_sd = Equation(m, "eq_su_sd", domain=[i, t], description="Startup-shutdown logic")
eq_su_sd[i, t].where[Ord(t) > 1] = (
    su[i, t] - sd[i, t] == u[i, t] - u[i, t.lag(1)]
)

# Minimum up time
eq_mut = Equation(m, "eq_mut", domain=[i, t], description="Minimum up time")
eq_mut[i, t].where[Ord(t) > 1] = (
    Sum(
        tt.where[(Ord(tt) >= Ord(t) - mut_p[i] + 1) & (Ord(tt) <= Ord(t))],
        su[i, tt]
    ) <= u[i, t]
)

# Minimum down time
eq_mdt = Equation(m, "eq_mdt", domain=[i, t], description="Minimum down time")
eq_mdt[i, t].where[Ord(t) > 1] = (
    Sum(
        tt.where[(Ord(tt) >= Ord(t) - mdt_p[i] + 1) & (Ord(tt) <= Ord(t))],
        sd[i, tt]
    ) <= 1 - u[i, t]
)

# Spinning reserve
eq_spin = Equation(m, "eq_spin", domain=t, description="Spinning reserve")
eq_spin[t] = (
    Sum(i, pmax_p[i] * u[i, t]) - Sum(i, P[i, t]) >= spin_p[t]
)

# Flow anti-symmetry
eq_antisym = Equation(m, "eq_antisym", domain=[z, zz, t], description="Flow anti-symmetry")
eq_antisym[z, zz, t].where[(tl_p[z, zz] > 0) & (tl_p[zz, z] > 0) & (Ord(z) < Ord(zz))] = (
    Flow[z, zz, t] + Flow[zz, z, t] == Number(0)
)

# ======================================================================
#  SECTION 6 - MODEL AND SOLVE
# ======================================================================

scuc_model = Model(
    m,
    name="SCUC_SCED",
    equations=m.getEquations(),
    problem=Problem.MIP,
    sense=Sense.MIN,
    objective=TotalCost,
)

print("Model built. Starting solve...")
print(f"  Generators: {len(GENERATORS)}  |  Time blocks: {len(BLOCKS)}  |  Zones: {len(ZONES)}")

from gamspy import Options
scuc_model.solve(
    options=Options(
        time_limit=7200,
        relative_optimality_gap=0.001,  # tightened: 0.1% gap
    ),
    output=None,
)

print(f"\nSolve complete.")
print(f"  Model status : {scuc_model.status}")
print(f"  Solver status: {scuc_model.solve_status}")
obj_val = float(TotalCost.records['level'].iloc[0])
print(f"  Objective    : Rs {obj_val:,.0f}")

# ======================================================================
#  SECTION 7 - RESULTS EXTRACTION
# ======================================================================

print("\nExtracting results...")

# Commitment schedule
u_df  = u.records[["i","t","level"]].rename(columns={"level":"committed"})
P_df  = P.records[["i","t","level"]].rename(columns={"level":"dispatch_MW"})
su_df = su.records[["i","t","level"]].rename(columns={"level":"startup"})
sd_df = sd.records[["i","t","level"]].rename(columns={"level":"shutdown"})

# Convert categorical columns to string (GAMSPy returns Categorical dtype)
for df in [u_df, P_df, su_df, sd_df]:
    df["i"] = df["i"].astype(str)
    df["t"] = df["t"].astype(str)

results = (u_df
    .merge(P_df,  on=["i","t"], how="left")
    .merge(su_df, on=["i","t"], how="left")
    .merge(sd_df, on=["i","t"], how="left")
)
# Ensure numeric columns
results["committed"]   = pd.to_numeric(results["committed"],   errors="coerce").fillna(0)
results["dispatch_MW"] = pd.to_numeric(results["dispatch_MW"], errors="coerce").fillna(0)
results["startup"]     = pd.to_numeric(results["startup"],     errors="coerce").fillna(0)
results["shutdown"]    = pd.to_numeric(results["shutdown"],    errors="coerce").fillna(0)

results["zone"]    = results["i"].map(GEN_ZONE)
results["vc_Rs"]   = results["i"].map(VC) * 250 * results["dispatch_MW"]
results["nl_Rs"]   = results["i"].map(NOLOAD) * 0.25 * results["committed"]
results["su_Rs"]   = results["i"].map(SU_COST) * results["startup"]
results["total_Rs"]= results["vc_Rs"] + results["nl_Rs"] + results["su_Rs"]

# Flow, curtailment and unserved energy
flow_df    = Flow.records[["z","zz","t","level"]].rename(columns={"level":"flow_MW"})
curtail_df = Curtail.records[["z","t","level"]].rename(columns={"level":"curtail_MW"})
use_df     = USE.records[["z","t","level"]].rename(columns={"level":"unserved_MW"})
for df in [use_df]:
    for col in df.columns:
        if df[col].dtype.name == "category":
            df[col] = df[col].astype(str)
use_df["unserved_MW"] = pd.to_numeric(use_df["unserved_MW"], errors="coerce").fillna(0)
total_use = use_df["unserved_MW"].sum() * 0.25
print(f"  Total Unserved Energy : {total_use:,.1f} MWh  (target: 0)")
# Detailed breakdown
nonzero_use = use_df[use_df["unserved_MW"] > 0.01].sort_values("unserved_MW", ascending=False)
if len(nonzero_use) > 0:
    print(f"  Unserved in {len(nonzero_use)} zone-blocks:")
    print(nonzero_use.groupby("z")["unserved_MW"].agg(["sum","count","max"]).round(2).to_string())
    print("  Top 5 worst:")
    print(nonzero_use.head(5)[["z","t","unserved_MW"]].to_string())
for df in [flow_df, curtail_df]:
    for col in df.columns:
        if df[col].dtype.name == "category":
            df[col] = df[col].astype(str)
flow_df["flow_MW"]       = pd.to_numeric(flow_df["flow_MW"],       errors="coerce").fillna(0)
curtail_df["curtail_MW"] = pd.to_numeric(curtail_df["curtail_MW"], errors="coerce").fillna(0)

# System summary per block
sys_summary = (results.groupby("t")
    .agg(total_gen_MW=("dispatch_MW","sum"),
         units_online=("committed","sum"),
         total_cost_Rs=("total_Rs","sum"))
    .reset_index()
)
sys_summary["demand_MW"]   = sys_summary["t"].map(DEMAND_T)
sys_summary["spin_req_MW"] = sys_summary["t"].map(SPIN_REQ)
sys_summary["spin_avail_MW"] = (
    results.groupby("t").apply(
        lambda df: (df["committed"] * df["i"].map(PMAX)).sum() - df["dispatch_MW"].sum()
    ).reset_index(drop=True)
)

# Unit summary
unit_summary = (results.groupby("i")
    .agg(total_MWh=("dispatch_MW", lambda x: x.sum()*0.25),
         num_startups=("startup","sum"),
         vc_cost_Rs=("vc_Rs","sum"),
         nl_cost_Rs=("nl_Rs","sum"),
         su_cost_Rs=("su_Rs","sum"),
         total_cost_Rs=("total_Rs","sum"))
    .reset_index()
)
unit_summary["zone"] = unit_summary["i"].map(GEN_ZONE)
unit_summary["pmax_MW"] = unit_summary["i"].map(PMAX)

# ======================================================================
#  SECTION 8 - EXPORT TO EXCEL
# ======================================================================

print("Writing results to Excel...")
out_file = "Gujarat_SCUC_Results.xlsx"

with pd.ExcelWriter(out_file, engine="openpyxl") as xw:
    # Sheet 1: Full dispatch schedule (pivot: generators x time blocks)
    pivot_dispatch = results.pivot(index="i", columns="t", values="dispatch_MW").fillna(0)
    pivot_dispatch.to_excel(xw, sheet_name="Dispatch_MW")

    # Sheet 2: Commitment schedule
    pivot_commit = results.pivot(index="i", columns="t", values="committed").fillna(0)
    pivot_commit.to_excel(xw, sheet_name="Commitment")

    # Sheet 3: Startup events
    pivot_su = results.pivot(index="i", columns="t", values="startup").fillna(0)
    pivot_su.to_excel(xw, sheet_name="Startups")

    # Sheet 4: System summary per block
    sys_summary.to_excel(xw, sheet_name="System_Summary", index=False)

    # Sheet 5: Unit cost summary
    unit_summary.sort_values("total_cost_Rs", ascending=False).to_excel(
        xw, sheet_name="Unit_Cost_Summary", index=False)

    # Sheet 6: Inter-zonal flows
    flow_df.to_excel(xw, sheet_name="Zonal_Flows", index=False)

    # Sheet 7: RE curtailment
    curtail_df.to_excel(xw, sheet_name="RE_Curtailment", index=False)

    # Sheet 8: Full flat results
    results.to_excel(xw, sheet_name="Full_Results", index=False)

    # Sheet 9: Unserved energy (should be zero after fix)
    use_df.to_excel(xw, sheet_name="Unserved_Energy", index=False)

print(f"Results written to {out_file}")
print("\n=== SOLUTION SUMMARY ===")
print(f"Total Cost          : Rs {obj_val:>20,.0f}")
print(f"Peak Demand         : {max(DEMAND_T.values()):>10,.0f} MW")
print(f"Units in model      : {len(GENERATORS):>10}")
print(f"Time blocks         : {len(BLOCKS):>10}")
print(f"Optimality gap      : <= 1%")
