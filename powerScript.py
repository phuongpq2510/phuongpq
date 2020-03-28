""" Replication of PowerScript library in Python
"""
from __future__ import print_function

__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2018, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.1.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

from pyOlrxAPI import *
from pyOlrxAPILib import *
import os,math
from datetime import datetime

#
def tiebranches(busCodeAZ,azNum):
    """
     Report tie branches by area or zone
        Args :
            busCodeAZ = BUS_nArea tie branches by AREA
                      = BUS_nZone tie branches by ZONE
            azNum : area/zone number (if <=0 all area/zone)

        Returns:
            string result

        Raises:
            OlrxAPIException
    """
    res  = "List of tie branches in the system"
    if busCodeAZ == BUS_nArea:
        if azNum<=0:
            res  += " (all area) "
        else:
            res  += " (with area = "+str(azNum) +")"
        res += "\nArea1   ,Area2   ,Branches"
    elif busCodeAZ == BUS_nZone:
        if azNum<=0:
            res  += " (all zone) "
        else:
            res  += " (with zone = "+str(azNum) +")"
        res += "\nZone1   ,Zone2   ,Branches"
    else:
        raise OlrxAPIException("Error busCodeAZ = BUS_nArea | BUS_nZone")

    #
    br,bus = getAllBranchesAndBus()
    # get result
    abr = {}
    for k in range(len(br)):
        b1 = bus[k]
        br1 = br[k]
        area = OlrxAPIGetEquipementData(b1,[busCodeAZ],VT_INTEGER)[0]
        #
        for i in range(len(b1)):
            for j in range(i+1, len(b1)):
                if  (area[i]!= area[j]) and (azNum<=0 or area[i]==azNum or area[j]==azNum):
                    abr[br1] = (min(area[i],area[j]),max(area[i],area[j]))
    # sort result
    sa = sorted(abr.items(), key=lambda kv: kv[1])
    # get result
    for v1 in sa:
        br1= v1[0]
        a1 = v1[1][0]
        a2 = v1[1][1]
        res += "\n"+str(a1).ljust(8)+","+str(a2).ljust(8)+","+OlrxAPIFullBranchName(br1)
    # print
    print(res)
    return res

def tielines(busCodeAZ,azNum):
    """
     Report tie Lines by area or zone
        Args :
            busCodeAZ = BUS_nArea tie branches by AREA
                      = BUS_nZone tie branches by ZONE
            azNum : area/zone number (if <=0 all area/zone)

        Returns:
            string result

        Raises:
            OlrxAPIException
    """
    res  = "List of tie lines in the system"
    if busCodeAZ == BUS_nArea:
        if azNum<=0:
            res  += " (all area) "
        else:
            res  += " (with area = "+str(azNum) +")"
        res += "\nArea1   ,Area2   ,Lines"
    elif busCodeAZ == BUS_nZone:
        if azNum<=0:
            res  += " (all zone) "
        else:
            res  += " (with zone = "+str(azNum) +")"
        res += "\nZone1   ,Zone2   ,Lines"
    else:
        raise OlrxAPIException("Error busCodeAZ = BUS_nArea | BUS_nZone")
    #
    bus = []
    br = []
    hndBr = c_int(0)
    while ( OLRXAPI_OK == OlrxAPIGetEquipment(c_int(TC_LINE), byref(hndBr) )) :
        bus.append(getBusByBranch(hndBr,TC_LINE))
        br.append(hndBr.value)

    # get result
    abr = {}
    for k in range(len(br)):
        b1  = bus[k]
        br1 = br[k]
        area = OlrxAPIGetEquipementData(b1,[busCodeAZ],VT_INTEGER)[0]
        #
        for i in range(len(b1)):
            for j in range(i+1, len(b1)):
                if  (area[i]!= area[j]) and (azNum<=0 or area[i]==azNum or area[j]==azNum):
                    abr[br1] = (min(area[i],area[j]),max(area[i],area[j]))
    # sort result
    sa = sorted(abr.items(), key=lambda kv: kv[1])
    # get result
    for v1 in sa:
        br1= v1[0]
        a1 = v1[1][0]
        a2 = v1[1][1]
        res += "\n"+str(a1).ljust(8)+","+str(a2).ljust(8)+","+OlrxAPIFullBranchName(br1)
    #print
    print(res)
    return res
#
def changebus(nParamCode,valOld,valNew):
    """
    Change all bus data
        Args :
            nParamCode : BUS_nArea/BUS_nZone
            valOld,valNew : valOld => valNew

        Returns:
            output: file OLR has changed => saved to0-=8n 9ij xx.OLR

        Raises:
            OlrxAPIException
    """
    # get all bus
    bhnd = c_int(0) #next cmd to seek the first bus
    kd = 0
    while ( OLRXAPI_OK == OlrxAPIGetEquipment(TC_BUS,byref(bhnd) ) ):
        v1 = c_int(0)
        if ( OLRXAPI_FAILURE == OlrxAPIGetData( bhnd, c_int(nParamCode), byref(v1) ) ) :
            raise OlrxAPIException(OlrxAPIErrorString())
        #
        if v1.value==valOld:
            kd+=1
            #change
            if OLRXAPI_OK != OlrxAPISetData( bhnd, c_int(nParamCode), byref(c_int(valNew)) ):
                raise OlrxAPIException(OlrxAPIErrorString())
            # Validation
            if OLRXAPI_OK != OlrxAPIPostData(bhnd):
                raise OlrxAPIException(OlrxAPIErrorString())
    #
    print("Change for "+str(kd) +" bus, valOld=", valOld,"valNew=", valNew)

    # save to xx.OLR
    olrFilePathNew = os.getcwd() + os.sep +"xx.OLR"
    if OLRXAPI_FAILURE == OlrxAPISaveDataFile(olrFilePathNew):
        raise OlrxAPIException(OlrxAPIErrorString())

#
def linez(hndLinePicked):
    """
    Report total line impedance and length
    Lines with tap buses are handled correctly
        Args :
            hndLinePicked : line picked
                            if <=0 all lines

        Returns:
            sres: string result
            kd: number of line found

        Raises:
            OlrxAPIException
    """
    testNotFound = True
    if hndLinePicked<=0:
        testNotFound = False
    #
    hndBrA = []
    hndBr1 = c_int(0)
    while ( OLRXAPI_OK == OlrxAPIGetEquipment(c_int(TC_LINE), byref(hndBr1) )) :
        hndBrA.append(hndBr1.value)
        if hndBr1.value==hndLinePicked:
            testNotFound = False
    #
    kd = 0
    sres  = str("No").ljust(6)+str(",Type   ")
    sres += "," + str("Bus1").ljust(30)+","+str("Bus2").ljust(30)+",ID  "
    sres += "," +"Z1".ljust(35)
    sres += "," +"Z0".ljust(35)
    sres += ",Length"
    if testNotFound:
        print("\nNumber of line found:",kd)
        return sres,kd
    #
    brOfTapBus = {}
    for hndBr in hndBrA:
        # bus
        bus  = getBusByBranch(hndBr,TC_LINE)
        nTap = OlrxAPIGetEquipementData(bus,[BUS_nTapBus],VT_INTEGER)[0] # if tap bus
        idBr = OlrxAPIGetEquipementData([hndBr],[LN_sID],VT_STRING)[0][0] # id branch
        test = False
        if hndLinePicked<=0 or hndLinePicked == hndBr: # all line or line=picked
            test = True
        #
        for i in range(len(bus)):
            bi = bus[i]
            if nTap[i]==1:
                try:
                    brOfTapBus[bi].append(hndBr)
                except:
                    brOfTapBus[bi] = []
                    brOfTapBus[bi].append(hndBr)
                #
                test = False
        # test= True => no Tap bus
        if test:
            typCode = [LN_dR,LN_dX,LN_dR0,LN_dX0,LN_dLength]
            val  = OlrxAPIGetEquipementData([hndBr],typCode,VT_DOUBLE)
            r1,x1,r0,x0,length = val[0][0],val[1][0],val[2][0],val[3][0],val[4][0]
            kV = OlrxAPIGetEquipementData([bus[0]],[BUS_dKVnominal],VT_DOUBLE)[0][0]
            #
            kd +=1
            sres +="\n" + str(kd).ljust(6) + ",Line   ,"
            sres += str(OlrxAPIFullBusName(bus[0])).ljust(30) + ","
            sres += str(OlrxAPIFullBusName(bus[1])).ljust(30) + ","
            sres += str(idBr).ljust(4) + ","
            sres += printImpedance(r1,x1,kV).ljust(35) + ","
            sres += printImpedance(r0,x0,kV).ljust(35) + ","
            sres +=  "{0:.2f}".format(length)
    # for line have many tap bus
    r = []
    for k in brOfTapBus.keys():
        r.append(set(brOfTapBus[k]))
    #
    for i in range(len(r)-1):
        for j in range(i+1,len(r)):
            if len(r[i])>0 and len(r[i])>0:
                vt = r[i] | r[j]
                if len(vt)<len(r[i]) + len(r[j]):
                    r[j]= vt
                    r[i]= set()
    # print segment
    for l1 in r:
        test = False
        if hndLinePicked<=0:
            test = True
        #
        for s1 in l1:
            if s1==hndLinePicked:
                test = True
        #
        if test and len(l1)>0:
            r1a,x1a,r0a,x0a,lengtha = 0,0,0,0,0
            busa = []
            for s1 in l1:
                bus = getBusByBranch(s1,TC_LINE)
                nTap = OlrxAPIGetEquipementData(bus,[BUS_nTapBus],VT_INTEGER)[0]
                for i in range(len(bus)):
                    if nTap[i]==0:
                        busa.append(bus[i])
                #
                typCode = [LN_dR,LN_dX,LN_dR0,LN_dX0,LN_dLength]
                val  = OlrxAPIGetEquipementData([s1],typCode,VT_DOUBLE)
                r1,x1,r0,x0,length = val[0][0],val[1][0],val[2][0],val[3][0],val[4][0]
                #
                sres +="\n" + str("").ljust(6) + ",Segment,"
                sres += str(OlrxAPIFullBusName(bus[0])).ljust(30) + ","
                sres += str(OlrxAPIFullBusName(bus[1])).ljust(30) + ","
                sres += str(idBr).ljust(4) + ","
                sres += printImpedance(r1,x1,kV).ljust(35) + ","
                sres += printImpedance(r0,x0,kV).ljust(35) + ","
                sres +=  "{0:.2f}".format(length)
                #
                r1a,x1a,r0a,x0a,lengtha = r1a+r1,x1a+x1,r0a+r0,x0a+x0,lengtha+length
            kd+=1
            sres +="\n" + str(kd).ljust(6) + ",Line   ,"
            if len(busa)==2:
                sres += str(OlrxAPIFullBusName(busa[0])).ljust(30) + ","
                sres += str(OlrxAPIFullBusName(busa[1])).ljust(30) + ","
                sres += str("").ljust(4) + ","
                sres += printImpedance(r1a,x1a,kV).ljust(35) + ","
                sres += printImpedance(r0a,x0a,kV).ljust(35) + ","
                sres +=  "{0:.2f}".format(lengtha)
            sres += "\n"
    print(sres)
    print("\nNumber of line found:",kd)
    return sres,kd

def networkutil(PickedHnd1):
    """
    Purpose: Find all remote ends of a relay group on transmission line.
    All taps are ignored. Close switches are included

    return:  number,s
    """
    PickedHnd = 4
    test = True
    hndRg = c_int(0)
    while ( OLRXAPI_OK == OlrxAPIGetEquipment(c_int(TC_RLYGROUP), byref(hndRg) )) :
        if hndRg.value==PickedHnd:
            test = False
    #
    if test:
        sres = "Please select a relay group on transmission line"
        print(sres)
        return 0,sres
    #
##    Call GetData( PickedHnd, RG_nBranchHnd, branchHnd& )
    hndBr = c_int(0)
    if ( OLRXAPI_FAILURE == OlrxAPIGetData( PickedHnd, c_int(RG_nBranchHnd), byref(hndBr) ) ) :
        raise OlrxAPIException(OlrxAPIErrorString())
    localBus = c_int(0)
    if ( OLRXAPI_FAILURE == OlrxAPIGetData( hndBr, c_int(BR_nBus1Hnd), byref(localBus) ) ) :
        raise OlrxAPIException(OlrxAPIErrorString())
    print ("\tlocal",OlrxAPIFullBusName(localBus))
    print ("\tbranch ",OlrxAPIFullBranchName(hndBr))
    #
##    typ = OlrxAPIEquipmentType(4)
##    print(typ)
    return


def main():
    # load orldll
    InitOlrxAPI("C:\\Program Files (x86)\\ASPEN\\1LPFv15\\\olxapi.dll")
    # open OLR file
    OlrxAPILoadDataFile(os.getcwd()+"\\SAMPLE30.OLR",1)  # TUAN09_A0    SAMPLE30
    #
    changebus(nParamCode=BUS_nArea,valOld=1,valNew=10)
    tiebranches(busCodeAZ=BUS_nArea,azNum=10)
##    tielines(busCodeAZ=BUS_nZone,azNum=10)
    linez(-4)
##    networkutil(4)


if __name__ == '__main__':
    t0 = datetime.now()
    main()
    print('\nRunTime: {}'.format(datetime.now() - t0))


