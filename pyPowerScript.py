"""
Demo usage of ASPEN OlrxAPI in Python.

transcipt some powerscript to python
1. changebus
2. tielies
3. linez
4. networkutil (remoteTerminal)


"""
from __future__ import print_function
from AspenPy import*
import sys,os

try:
    import OlxAPILib
    import OlxAPI
    import OlxAPIConst
except:
    pass



def changebus(nParamCode,valOld,valNew):
    """
    Change all bus data
        Args :
            nParamCode : BUS_nArea/BUS_nZone
            valOld,valNew : valOld => valNew

        Returns:
            kd: number of bus changed

        Raises:
            OlxAPI.Exception
    """
    kd = 0
    if valOld==valNew:
        return kd
    # get all bus
    bhnd = c_int(0) #next cmd to seek the first bus
    while ( OLRXAPI_OK == OlxAPI.GetEquipment(TC_BUS,byref(bhnd) ) ):
        v1 = c_int(0)
        if ( OLRXAPI_FAILURE ==OlxAPI.GetData( bhnd, c_int(nParamCode), byref(v1) ) ) :
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        #
        if v1.value==valOld:
            kd+=1
            #change
            if OLRXAPI_OK != OlxAPI.SetData( bhnd, c_int(nParamCode), byref(c_int(valNew)) ):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            # Validation
            if OLRXAPI_OK !=  OlxAPI.PostData(bhnd):
                 raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    return kd

def tielines(nParamCode,azNum,equiCode):
    """
     Report tie branches by area or zone
        Args :
            nParamCode = BUS_nArea tie lines by AREA
                       = BUS_nZone tie lines by ZONE
            azNum : area/zone number (if <=0 all area/zone)
            equiCode: code of equipement
               example:
                    = [TC_LINE] => report tie lines
                    = [NULL]    => report tie branches
                    = [TC_LINE, TC_SCAP] => report tie lines and Serie capacitor/reactor

        Returns:
            nC : number of tie branches
            sres: string result

        Raises:
            OlrxAPIException
    """
    equiCode1 = equiCode
    if len(equiCode1)==0:
        equiCode1 = [TC_LINE, TC_SCAP, TC_PS, TC_SWITCH, TC_XFMR, TC_XFMR3]
    #
    abr = {}
    for ec1 in equiCode1:
        eqa = OlxAPILib.getEquipementHandle(ec1)
        for ehnd1 in eqa:
            bus = OlxAPILib.getBusByEquipement(ehnd1,ec1)
            az = OlxAPILib.getEquipementData(bus,[nParamCode],VT_INTEGER)[0]
            #
            for i in range(len(bus)):
                for j in range(i+1, len(bus)):
                    if  (az[i]!= az[j]) and (azNum<=0 or az[i]==azNum or az[j]==azNum):
                        abr[ehnd1] = (min(az[i],az[j]),max(az[i],az[j]))
    # sort result
    sa = sorted(abr.items(), key=lambda kv: kv[1])
    # get string result
    sres = ""
    for v1 in sa:
        ec1= v1[0]
        a1 = v1[1][0]
        a2 = v1[1][1]
        sres += str(a1).ljust(8)+","+str(a2).ljust(8)+"," + OlxAPI.FullBranchName(ec1)+"\n"
    nC = len(sa)
    return nC,sres


def linez_fromBus(bhnd):
    """
     Report total line impedance and length
        All line from a bus
        All taps are ignored. Close switches are included

        Args :
            bhnd = bus handle

        Returns:
            nC : number of lines
            sres: string result

        Raises:
            OlrxAPIException
    """
    typeConsi = [TC_LINE,TC_SWITCH,TC_SCAP]
    allBr = OlxAPILib.getBusEquipmentData([bhnd],[TC_BRANCH],VT_INTEGER)[0][0]
    sres = ""
    sres += str("No").ljust(6)+str(",Type   ").ljust(15)
    sres += "," + str("Bus1").ljust(30)+","+str("Bus2").ljust(30)+",ID  "
    sres += "," +"Z1".ljust(35)
    sres += "," +"Z0".ljust(35)
    sres += ",Length"
    #
    kd = 0
    for hndBr in allBr:
        if OlxAPILib.branchIsInType(hndBr,typeConsi):
            bra = OlxAPILib.lineComponents(hndBr,typeConsi)
            for bra1 in bra:
                kd +=1
                s1 = __getsImpedanceLine(kd,bra1)
                sres +=s1
    return kd,sres

def getRemoteTerminals(bhnd):
    """
    Purpose: Find all remote end of all branches (from a bus)
             All taps are ignored. Close switches are included
             =>a test function of OlxAPILib.getRemoteTerminals(hndBr)

        Args :
            bhnd :  bus handle

        returns :
            bus_res [] list of terminal bus

        Raises:
            OlrxAPIException
    """
    br_res = OlxAPILib.getBusEquipmentData([bhnd],[TC_BRANCH],VT_INTEGER)[0][0]
    bus_res = []
    #
    for br1 in br_res:
        ba = OlxAPILib.getRemoteTerminals(br1)
        bus_res.append(ba)
    return br_res,bus_res

def __getsImpedanceLine(kd,bra): #[TC_LINE,TC_SWITCH,TC_SCAP]
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

def __test_changebus():
    olrFile="C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python\\Sample30.OLR"
    setPara_sys_argv(idx=1,val=olrFile)
    #
    load_olxapi_dll()
    open_olrFile(readonly=0)
    #
    kb = changebus(BUS_nArea,valOld=1,valNew=10)
    print(str(kb) +" bus have changed area ("+ str(1) + "=>"+ str(10) + ")")
    #
    kb = changebus(BUS_nZone,valOld=1,valNew=10)
    print(str(kb) +" bus have changed zone ("+ str(1) + "=>"+ str(10) + ")")


def __test_tielines():
    olrFile="C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python\\Sample30_1.OLR"
    setPara_sys_argv(idx=1,val=olrFile)
    #
    load_olxapi_dll()
    open_olrFile(readonly=0)
    nC,sres = tielines(nParamCode=BUS_nArea,azNum=-1,equiCode=[TC_LINE]) #Tie lines by area
    print(nC," tie lines in the system (all Area)")
    print("Area1   ,Area2   ,Lines")
    print(sres)
    #
    nC,sres = tielines(nParamCode=BUS_nZone,azNum=1,equiCode=[]) #tie Branches by zone
    print(nC," tie branches in the system (zone=1)")
    print("Zone1   ,Zone2   ,Branches")
    print(sres)

def __test_linez_fromBus():
    olrFile="C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python\\Sample30_2.OLR"
    setPara_sys_argv(idx=1,val=olrFile)
    load_olxapi_dll()
    open_olrFile(readonly=0)
    #
    bhnd  = OlxAPI.FindBus("KENTUCKY",33.0) # "KENTUCKY",33.0 "NEVADA",132.0
    kd,sres = linez_fromBus(bhnd)
    print(str(kd)+" Line impedances from bus: " + OlxAPI.FullBusName(bhnd))
    print(sres)

def __test_getRemoteTerminal():
    olrFile="C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python\\Sample30_2.OLR"
    setPara_sys_argv(idx=1,val=olrFile)
    load_olxapi_dll()
    open_olrFile(readonly=0)
    #
    bhnd  = OlxAPI.FindBus("KENTUCKY",33.0) # "KENTUCKY",33.0 "NEVADA",132.0
    br_res,bus_res = getRemoteTerminals(bhnd)
    #
    print(str(len(br_res))+" branches from bus: " + OlxAPI.FullBusName(bhnd))
    for i in range(len(br_res)):
        print(str(i+1)+OlxAPI.FullBranchName(br_res[i]))
        print("\t"+str(len(bus_res[i]))+ " remote terminals:")
        for j in range(len(bus_res[i])):
            print("\t\t"+str(j+1)+" "+OlxAPI.FullBusName(bus_res[i][j]))

# run with input arguments
if __name__ == '__main__':
##    __test_changebus()
##    __test_tielines()
##    __test_linez_fromBus()
    __test_getRemoteTerminal()








