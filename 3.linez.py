"""
Demo usage of ASPEN OlrxAPI in Python.

Purpose:
    Report total line impedance and length
    Lines with tap buses are handled correctly

"""
from __future__ import print_function
from AspenPy import*
try:
    import OlxAPILib
    import OlxAPI
    import OlxAPIConst
except:
    pass

__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.1.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

import Tkinter as tk
import tkFileDialog,tkMessageBox,ttk
import sys,os

#
def getsImpedanceLine(kd,bra): #[TC_LINE,TC_SWITCH,TC_SCAP]
    sres = ""
    typa =  OlxAPILib.getEquipementData(bra,[BR_nType],VT_INTEGER)[0]
    #
    bus = OlxAPILib.getEquipementData([bra[0],bra[len(bra)-1]],[BR_nBus1Hnd],VT_INTEGER)
    b1,b2 = bus[0][0],bus[0][1]
    #
    kV = OlxAPILib.getEquipementData([b1],[BUS_dKVnominal],VT_DOUBLE)[0][0]
    #
    if len(bra)==2: # simple line
        br1 = bra[0]
        typ1 = typa[0]
        equi_hnd = OlxAPILib.getEquipementData([br1],[BR_nHandle],VT_INTEGER)[0][0]
        #
        if typ1== TC_LINE:
            val  = OlxAPILib.getEquipementData([equi_hnd], [LN_dR,LN_dX,LN_dR0,LN_dX0,LN_dLength] ,VT_DOUBLE)
            r1,x1,r0,x0,length = val[0][0],val[1][0],val[2][0],val[3][0],val[4][0]
            #
            idBr = OlxAPILib.getEquipementData([equi_hnd],[LN_sID],VT_STRING)[0][0] # id line
            #
            sres +="\n" + str(kd).ljust(5) + str(",Line").ljust(15) + ","
            sres += str(OlxAPI.FullBusName(b1)).ljust(30) + ","
            sres += str(OlxAPI.FullBusName(b2)).ljust(30) + ","
            sres += str(idBr).ljust(4) + ","
            sres += OlxAPILib.printImpedance(r1,x1,kV).ljust(35) + ","
            sres += OlxAPILib.printImpedance(r0,x0,kV).ljust(35) + ","
            sres +=  "{0:.2f}".format(length)
        elif typ1 == TC_SWITCH:
            idBr = OlxAPILib.getEquipementData([equi_hnd],[SW_sID],VT_STRING)[0][0] # id switch
            #
            sres +="\n" + str(kd).ljust(5) + str(",Switch").ljust(15) + ","
            sres += str(OlxAPI.FullBusName(b1)).ljust(30) + ","
            sres += str(OlxAPI.FullBusName(b2)).ljust(30) + ","
            sres += str(idBr).ljust(4) + ","
        elif typ1 == TC_SCAP:
            val  = OlxAPILib.getEquipementData([equi_hnd], [SC_dR,SC_dX,SC_dR0,SC_dX0] ,VT_DOUBLE)
            r1,x1,r0,x0 = val[0][0],val[1][0],val[2][0],val[3][0]
            #
            idBr = OlxAPILib.getEquipementData([equi_hnd],[SC_sID],VT_STRING)[0][0] # id switch
            #
            sres +="\n" + str(kd).ljust(5) + str(",SerieCap").ljust(15) + ","
            sres += str(OlxAPI.FullBusName(b1)).ljust(30) + ","
            sres += str(OlxAPI.FullBusName(b2)).ljust(30) + ","
            sres += str(idBr).ljust(4) + ","
            sres += OlxAPILib.printImpedance(r1,x1,kV).ljust(35) + ","
            sres += OlxAPILib.printImpedance(r0,x0,kV).ljust(35) + ","
    else: #Line with Tap bus
        r1a,x1a,r0a,x0a,lengtha = 0,0,0,0,0
        for i in range(len(bra)-1):
            br1 = bra[i]
            typ1 = typa[i]
            bus = OlxAPILib.getEquipementData([br1],[BR_nBus1Hnd,BR_nBus2Hnd],VT_INTEGER)
            b1i,b2i = bus[0][0],bus[1][0]
            #
            equi_hnd = OlxAPILib.getEquipementData([br1],[BR_nHandle],VT_INTEGER)[0][0]
            if typ1== TC_LINE:
                val  = OlxAPILib.getEquipementData([equi_hnd], [LN_dR,LN_dX,LN_dR0,LN_dX0,LN_dLength] ,VT_DOUBLE)
                r1,x1,r0,x0,length = val[0][0],val[1][0],val[2][0],val[3][0],val[4][0]
                r1a,x1a,r0a,x0a,lengtha = r1a+r1,x1a+x1,r0a+r0,x0a+x0,lengtha+length
                #
                idBr = OlxAPILib.getEquipementData([equi_hnd],[LN_sID],VT_STRING)[0][0] # id line
                #
                sres +="\n" + str("").ljust(5) + str(",  Line_seg").ljust(15) + ","
                sres += str(OlxAPI.FullBusName(b1i)).ljust(30) + ","
                sres += str(OlxAPI.FullBusName(b2i)).ljust(30) + ","
                sres += str(idBr).ljust(4) + ","
                sres += OlxAPILib.printImpedance(r1,x1,kV).ljust(35) + ","
                sres += OlxAPILib.printImpedance(r0,x0,kV).ljust(35) + ","
                sres +=  "{0:.2f}".format(length)
            elif typ1 == TC_SWITCH:
                idBr = OlxAPILib.getEquipementData([equi_hnd],[SW_sID],VT_STRING)[0][0] # id switch
                #
                sres +="\n" + str("").ljust(5) + str(",  Switch_seg").ljust(15) + ","
                sres += str(OlxAPI.FullBusName(b1i)).ljust(30) + ","
                sres += str(OlxAPI.FullBusName(b2i)).ljust(30) + ","
                sres += str(idBr).ljust(4) + ","
            elif typ1 == TC_SCAP:
                val  = OlxAPILib.getEquipementData([equi_hnd], [SC_dR,SC_dX,SC_dR0,SC_dX0] ,VT_DOUBLE)
                r1,x1,r0,x0 = val[0][0],val[1][0],val[2][0],val[3][0]
                r1a,x1a,r0a,x0a = r1a+r1,x1a+x1,r0a+r0,x0a+x0
                #
                idBr = OlxAPILib.getEquipementData([equi_hnd],[SC_sID],VT_STRING)[0][0] # id switch
                #
                sres +="\n" + str("").ljust(5) + str(",  SerieCap_seg").ljust(15) + ","
                sres += str(OlxAPI.FullBusName(b1i)).ljust(30) + ","
                sres += str(OlxAPI.FullBusName(b2i)).ljust(30) + ","
                sres += str(idBr).ljust(4) + ","
                sres += OlxAPILib.printImpedance(r1,x1,kV).ljust(35) + ","
                sres += OlxAPILib.printImpedance(r0,x0,kV).ljust(35) + ","
        #summary
        sres +="\n" + str(kd).ljust(5) + str(",Line_sum").ljust(15) + ","
        sres += str(OlxAPI.FullBusName(b1)).ljust(30) + ","
        sres += str(OlxAPI.FullBusName(b2)).ljust(30) + ","
        sres += str("").ljust(4) + ","
        sres += OlxAPILib.printImpedance(r1a,x1a,kV).ljust(35) + ","
        sres += OlxAPILib.printImpedance(r0a,x0a,kV).ljust(35) + ","
        sres +=  "{0:.2f}".format(lengtha)
    return sres +"\n"
##
def linez_fromBus(bsName,bsKV):
    bhnd  = OlxAPI.FindBus(bsName, bsKV)
    allBr = OlxAPILib.getBusEquipmentData([bhnd],[TC_BRANCH],VT_INTEGER)[0][0]
    sres = "All lines from bus: " + OlxAPI.FullBusName(bhnd)
    sres += str("\nNo").ljust(6)+str(",Type   ").ljust(15)
    sres += "," + str("Bus1").ljust(30)+","+str("Bus2").ljust(30)+",ID  "
    sres += "," +"Z1".ljust(35)
    sres += "," +"Z0".ljust(35)
    sres += ",Length"
    #
    kd = 1
    for hndBr in allBr:
        if OlxAPILib.branchIsLineComponent(hndBr):
            bra = OlxAPILib.lineComponents(hndBr)
            for bra1 in bra:
                s1 = getsImpedanceLine(kd,bra1)
                sres +=s1
                kd +=1
    print(sres)
    return sres

# run with input arguments
if __name__ == '__main__':
    """
    Run with input arguments
        sys.argv[1] = (str) OLR file
        sys.argv[2] = (Int 1/0) loaded DLL
    """
    # test and set default input (if need)
    try:
        olrFile = sys.argv[1]
    except:
        olr = OLXAPI_PYTHON_PATH + "\\SAMPLE30_1.OLR" # OLR_DEFAULT_PATH
        setPara_sys_argv(idx=1,val=olr)
    olrFile = sys.argv[1]
    #
    load_olxapi_dll()
    open_olrFile(readonly=0)
    ##
##    linez_fromBus(bsName="ROANOKE", bsKV=13.8) #FIELDALE 132  KENTUCKY 33 ROANOKE 13.8
##    linez_fromBus(bsName="FIELDALE", bsKV=132.0)
    linez_fromBus(bsName="KENTUCKY", bsKV=33.0)
##    linez_fromBus(bsName="ARIZONA", bsKV=132.0)
##    linez_fromBus(bsName="COLORADO", bsKV=33.0)






