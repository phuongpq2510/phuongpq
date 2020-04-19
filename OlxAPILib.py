"""
Demo usage of ASPEN OlxAPI in Python.
common utils
"""
from __future__ import print_function

__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.2.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

from ctypes import *
from OlxAPIConst import *
import OlxAPI
import math
import os,time
# Tkinter python 2|3
try:
    import tkinter as tk
    import tkinter.filedialog as tkf
    import tkinter.messagebox as tkm
    from tkinter import ttk
except:
    import Tkinter as tk
    import tkFileDialog as tkf
    import tkMessageBox as tkm
    import ttk

#
def load_olxapi_dll():
    """
    Load olxapi.dll with test in
        ARGVS.ld = (Int 1/0) loaded DLL
    """
    if ARGVS.ld == None:
        ARGVS.ld = 0
    loadDLL = int(ARGVS.ld)
    #
    if loadDLL==0:
        OlxAPI.InitOlxAPI(OLXAPI_DLL_PATH)
        ARGVS.ld = 1
        print ("\tOlxAPI : Version:"+str(OlxAPI.Version())+" Build: "+str(OlxAPI.BuildNumber()))
#
def open_olrFile(readonly):
    """
    Open OLR file in
        ARGVS.fi = (str) OLR file
            if NONE OLR_DEFAULT_SAMPLE
    Args:
        readonly (int): open in read-only mode. 1-true; 0-false
    """
    if ARGVS.fi == None:
        ARGVS.fi = OLR_DEFAULT_SAMPLE
    olrFile = str(ARGVS.fi)
    #
    if OLRXAPI_OK == OlxAPI.LoadDataFile(olrFile,readonly):
         print("\tFile opened successfully: " + str(olrFile))
    else:
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())

def get_equipment(args):
    """Get OLR network object handle
    """
    c_tc = c_int(args["tc"])
    c_hnd = c_int(args["hnd"])
    ret = OlxAPI.GetEquipment(c_tc, pointer(c_hnd) )
    if OLRXAPI_OK == ret:
       args["hnd"] = c_hnd.value
    return ret

def set_data(args):
    """Set network object data field
    """
    token = args["token"]
    hnd = args["hnd"]
    vt = token//100
    if vt == VT_STRING:
        c_data = c_char_p(args["data"])
    elif vt == VT_DOUBLE:
        c_data =c_double(args["data"])
    elif vt == VT_INTEGER:
        c_data = c_int(args["data"])
    else:
        tc = OlxAPI.EquipmentType(hnd)
        if tc == TC_GENUNIT and (token == GU_vdR or token == GU_vdX):
            count = 5
        elif tc == TC_LOADUNIT and (token == LU_vdMW or token == LU_vdMVAR):
            count = 3
        elif tc == TC_LINE and token == LN_vdRating:
            count = 4
        array = args["data"]
        c_data = (c_double*len(array))(*array)
    return OlxAPI.SetData(c_int(hnd), c_int(token), byref(c_data))

def post_data(args):
    """Post network object data
    """
    return OlxAPI.PostData(c_int(args["hnd"]))

c_GetDataBuf = create_string_buffer(b'\000' * 10*1024*1024)    # 10 KB buffer for string data
c_GetDataBuf_double = c_double(0)
c_GetDataBuf_int = c_int(0)
def get_data(args):
    """Get network object data field value
    """
    c_token = c_int(args["token"])
    c_hnd = c_int(args["hnd"])
    try:
        data = args["data"]
    except:
        data = None
    dataBuf = make_GetDataBuf( c_token, data )
    ret = OlxAPI.GetData(c_hnd, c_token, byref(dataBuf))
    if OLRXAPI_OK == ret:
        args["data"] = process_GetDataBuf(dataBuf,c_token,c_hnd)
    return ret

def make_GetDataBuf(token,data):
    """Prepare correct data buffer for OlxAPI.GetData()
    """
    global c_GetDataBuf
    global c_GetDataBuf_double
    global c_GetDataBuf_int
    vt = token.value//100
    if vt == VT_DOUBLE:
        if data != None:
            try:
                c_GetDataBuf_double = c_double(data)
            except:
                pass
        return c_GetDataBuf_double
    elif vt == VT_INTEGER:
        if data != None:
            try:
                c_GetDataBuf_int = c_int(data)
            except:
                pass
        return c_GetDataBuf_int
    else:
        try:
            c_GetDataBuf.value = data
        except:
            pass
        return c_GetDataBuf

def process_GetDataBuf(buf,token,hnd):
    """Convert GetData binary data buffer into Python object of correct type
    """
    vt = token.value//100
    if vt == VT_STRING:
        return buf.value #
    elif vt == VT_DOUBLE:
        return buf.value
    elif vt == VT_INTEGER:
        return buf.value
    else:
        array = []
        tc = OlxAPI.EquipmentType(hnd)
        if tc == TC_BREAKER and (token.value == BK_vnG1DevHnd or \
            token.value == BK_vnG2DevHnd or \
            token.value == BK_vnG1OutageHnd or \
            token.value == BK_vnG2OutageHnd):
            val = cast(buf, POINTER(c_int*MXSBKF)).contents  # int array of size MXSBKF
            for ii in range(0,MXSBKF-1):
                array.append(val[ii])
                if array[ii] == 0:
                    break
        elif (tc == TC_SVD and (token.value == SV_vnNoStep)):
            val = cast(buf, POINTER(c_int*8)).contents  # int array of size 8
            for ii in range(0,7):
                array.append(val[ii])
        elif (tc == TC_RLYDSP and (token.value == DP_vParams or token.value == DP_vParamLabels)) or \
             (tc == TC_RLYDSG and (token.value == DG_vParams or token.value == DG_vParamLabels)):
            # String with tab delimited fields
            return (cast(buf, c_char_p).value).split("\t")
        else:
            if tc == TC_GENUNIT and (token.value == GU_vdR or token.value == GU_vdX):
                count = 5
            elif tc == TC_LOADUNIT and (token.value == LU_vdMW or token.value == LU_vdMVAR):
                count = 3
            elif tc == TC_SVD and (token.value == SV_vdBinc or token.value == SV_vdB0inc):
                count = 3
            elif tc == TC_LINE and token.value == LN_vdRating:
                count = 4
            elif tc == TC_RLYGROUP and token.value == RG_vdRecloseInt:
                count = 3
            elif tc == TC_RLYOCG and token.value == OG_vdDirSetting:
                count = 2
            elif tc == TC_RLYOCP and token.value == OP_vdDirSetting:
                count = 2
            elif tc == TC_RLYDSG and token.value == DG_vdParams:
                count = MXDSPARAMS
            elif tc == TC_RLYDSG and (token.value == DG_vdDelay or token.value == DG_vdReach or token.value == DG_vdReach1):
                count = MXZONE
            elif tc == TC_RLYDSP and token.value == DP_vParams:
                count = MXDSPARAMS
            elif tc == TC_RLYDSP and (token.value == DP_vdDelay or token.value == DP_vdReach or token.value == DP_vdReach1):
                count = MXZONE
            elif tc == TC_CCGEN and (token.value == CC_vdV or token.value == CC_vdI or token.value == CC_vdAng):
                count = MAXCCV
            elif tc == TC_BREAKER and (token.value == BK_vdRecloseInt1 or token.value == BK_vdRecloseInt2):
                count = 3
            val = cast(buf, POINTER(c_double*count)).contents  # double array of size count
            for v in val:
                array.append(v)
        return array

def run_busFault(busHnd):
    """Report fault on a bus
    """
    hnd = c_int(busHnd)
    fltConn = (c_int*4)(0,0,1,0)   # 3LG, 2LG, 1LG, LL
    fltOpt = (c_double*15)(0)
    fltOpt[0] = 1       # Bus or Close-in
    fltOpt[1] = 0       # Bus or Close-in w/ outage
    fltOpt[2] = 0       # Bus or Close-in with end opened
    fltOpt[3] = 0       # Bus or Close#-n with end opened w/ outage
    fltOpt[4] = 0       # Remote bus
    fltOpt[5] = 0       # Remote bus w/ outage
    fltOpt[6] = 0       # Line end
    fltOpt[7] = 0       # Line end w/ outage
    fltOpt[8] = 0       # Intermediate %
    fltOpt[9] = 0       # Intermediate % w/ outage
    fltOpt[10] = 0      # Intermediate % with end opened
    fltOpt[11] = 0      # Intermediate % with end opened w/ outage
    fltOpt[12] = 0      # Auto seq. Intermediate % from [*] = 0
    fltOpt[13] = 0      # Auto seq. Intermediate % to [*] = 0
    fltOpt[14] = 0      # Outage line grounding admittance in mho [***] = 0.
    outageLst = (c_int*100)(0)
    outageLst[0] = 0
    outageOpt = (c_int*4)(0)
    outageOpt[0] = 0
    fltR = c_double(0.0)
    fltX = c_double(0.0)
    clearPrev = c_int(1)
    if OLRXAPI_FAILURE == OlxAPI.DoFault(hnd, fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev):
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    if OLRXAPI_FAILURE == OlxAPI.PickFault(c_int(SF_FIRST),c_int(9)):
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    print ("======================")
    print (OlxAPI.FaultDescription(0))
    hndBr = c_int(0);
    hndBus2 = c_int(0)
    vd12Mag = (c_double*12)(0.0)
    vd12Ang = (c_double*12)(0.0)
    vd9Mag = (c_double*9)(0.0)
    vd9Ang = (c_double*9)(0.0)
    vd3Mag = (c_double*3)(0.0)
    vd3Ang = (c_double*3)(0.0)
    while ( OLRXAPI_OK == OlxAPI.GetBusEquipment( hnd, c_int(TC_BRANCH), byref(hndBr) ) ) :
        if ( OLRXAPI_FAILURE == OlxAPI.GetData( hndBr, c_int(BR_nBus2Hnd), byref(hndBus2) ) ) :
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if ( OLRXAPI_FAILURE == OlxAPI.GetPSCVoltage( hndBr, vd3Mag, vd3Ang, c_int(2) ) ) :
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if ( OLRXAPI_FAILURE == OlxAPI.GetSCVoltage( hndBr, vd9Mag, vd9Ang, c_int(4) ) ) :
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        # Voltage on bus 1
        print (OlxAPI.FullBusName( hnd ), \
              "VP (pu)=", vd3Mag[0], "@", vd3Ang[0],    \
              "Va=", vd9Mag[0], "@", vd9Ang[0],    \
              "Vb=", vd9Mag[1], "@", vd9Ang[1],    \
              "Vc=", vd9Mag[2], "@", vd9Ang[2])
        # Voltage on bus 2
        print (OlxAPI.FullBusName( hndBus2 ), \
              "VP (pu)=", vd3Mag[1], "@", vd3Ang[1],    \
              "Va=", vd9Mag[3], "@", vd9Ang[3],    \
              "Vb=", vd9Mag[4], "@", vd9Ang[4],    \
              "Vc=", vd9Mag[5], "@", vd9Ang[5])

        if ( OLRXAPI_FAILURE == OlxAPI.GetSCCurrent( hndBr, vd12Mag, vd12Ang, 4 ) ) :
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        # Current from 1
        print ("Ia=", vd12Mag[0], "@", vd12Ang[0],    \
              "Ib=", vd12Mag[1], "@", vd12Ang[1],    \
              "Ic=", vd12Mag[2], "@", vd12Ang[2])
        # Relay time
        hndRlyGroup = c_int(0)
        if ( OLRXAPI_OK == OlxAPI.GetData( hndBr, c_int(BR_nRlyGrp1Hnd), byref(hndRlyGroup) ) ) :
            if hndRlyGroup != 0 :
                print_relayTime(hndRlyGroup)
        if ( OLRXAPI_OK == OlxAPI.GetData( hndBr, c_int(BR_nRlyGrp2Hnd), byref(hndRlyGroup) ) ) :
            if hndRlyGroup != 0 :
                print_relayTime(hndRlyGroup)

def print_relayTime(hndRlyGroup):
    """Print operating time of all relays in a relay group
    """
    hndRelay = c_int(0)
    while OLRXAPI_OK == OlxAPI.GetRelay( hndRlyGroup, byref(hndRelay) ) :
        print (OlxAPI.FullRelayName( hndRelay ))
        dTime = c_double(0)
        szDevice = create_string_buffer(b'\000' * 128)
        flag = c_int(0)
        if ( OLRXAPI_FAILURE == OlxAPI.GetRelayTime( hndRelay, c_double(1.0), flag, byref(dTime), szDevice ) ):
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        print (" time (s) = " + str(dTime.value) + "  device= " + str(szDevice.value))

def run_steppedEvent(busHnd):
    """Run stepped-event simulation on a bus
    """
    hnd = c_int(busHnd)
    runOpt = (c_int*5)(1,1,1,1,1)   # OCG, OCP, DSG, DSP, SCHEME
    fltOpt = (c_double*64)(0)
    fltOpt[0] = 1       #    Fault connection code
                        #    1=3LG
                        #    2=2LG BC, 3=2LG CA, 4=2LG AB
                        #    5=1LG A, 5=1LG B, 6=1LG C
                        #    7=LL BC, 7=LL CA, 8=LL AB

    fltOpt[1] = 0       # Intermediate percent between 0.01-99.99. 0 for a close-in fault.
                        #This parameter is ignored if nDevHnd is a bus handle.
    fltOpt[2] = 0       # Fault resistance, ohm
    fltOpt[3] = 0       # Fault reactance, ohm
    fltOpt[4] = 0       # Zero or Fault connection code for additional user event
    fltOpt[4+1] = 0     # Time  of additional user event, seconds.
    fltOpt[4+2] = 0     # Fault resistance in additinoal user event, ohm
    fltOpt[4+3] = 0     # Fault reactancein additinoal user event, ohm
                        #....
    noTiers = c_int(5)
    if OLRXAPI_FAILURE == OlxAPI.DoSteppedEvent(hnd, fltOpt, runOpt, noTiers):
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    # Call GetSteppedEvent with 0 to get total number of events simulated
    dTime = c_double(0)
    dCurrent = c_double(0)
    nUserEvent = c_int(0)
    szEventDesc = create_string_buffer(b'\000' * 512 * 4)     # 4*512 bytes buffer for event description
    szFaultDest = create_string_buffer(b'\000' * 512 * 50)    # 50*512 bytes buffer for fault description
    nSteps = OlxAPI.GetSteppedEvent( c_int(0), byref(dTime), byref(dCurrent),
                                               byref(nUserEvent), szEventDesc, szFaultDest )
    print ("Stepped-event simulation completed successfully with ", nSteps-1, " events")
    for ii in range(1, nSteps):
        OlxAPI.GetSteppedEvent( c_int(ii), byref(dTime), byref(dCurrent),
                                          byref(nUserEvent), szEventDesc, szFaultDest )
        print ("Time = ", dTime.value, " Current= ", dCurrent.value)
        print (cast(szFaultDest, c_char_p).value)
        print (cast(szEventDesc, c_char_p).value)

def branchSearch( bsBusName1, bsKV1, bsBusName2, bsKV2, sCKID ):
    hnd1    = OlxAPI.FindBus( bsBusName1, bsKV1 )
    if hnd1 == OLRXAPI_FAILURE:
            print ("Bus ", bsBusName1, bsKV1, " not found")
    hnd2    = OlxAPI.FindBus( bsBusName2, bsKV2 )
    if hnd2 == OLRXAPI_FAILURE:
            print ("Bus ", bsBusName2, bsKV2, " not found")

    hndBr = c_int(0)
    while ( OLRXAPI_OK == OlxAPI.GetBusEquipment( hnd1, c_int(TC_BRANCH), byref(hndBr) ) ) :
        argsGetData = {}
        argsGetData["hnd"] = hndBr.value
        argsGetData["token"] = BR_nBus2Hnd
        if (OLRXAPI_OK == get_data(argsGetData)):
            hndFarBus = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if hndFarBus == hnd2:
            argsGetData = {}
            argsGetData["hnd"] = hndBr.value
            argsGetData["token"] = BR_nType
            if (OLRXAPI_OK == get_data(argsGetData)):
                brType = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
            if brType == TC_LINE:
                typeCode = LN_sID
            elif brType == TC_XFMR:
                typeCode = XF_sID
            elif brType == TC_XFMR3:
                typeCode = X3_sID
            elif brType == TC_PS:
                typeCode = PS_sID
            argsGetData = {}
            argsGetData["hnd"] = hndBr.value
            argsGetData["token"] = BR_nHandle
            if (OLRXAPI_OK == get_data(argsGetData)):
                itemHnd = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
            argsGetData = {}
            argsGetData["hnd"] = itemHnd
            argsGetData["token"] = typeCode
            if (OLRXAPI_OK == get_data(argsGetData)):
                sID = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
            if sID == sCKID:
                return hndBr.value

def binarySearch(Array, nKey, nMin, nMax):
    while (nMax >= nMin):
        nMid = (nMax + nMin) / 2
        if Array[nMid] == nKey:
            return nMid
        else:
            if Array[nMid] < nKey:
                nMin = nMid + 1
            else:
                nMax = nMid - 1
    return -1

def compuOneLiner(nLineHnd, ProcessedHnd, hndOffset):
    argsGetData = {}
    argsGetData["hnd"] = nLineHnd
    argsGetData["token"] = LN_nBus1Hnd
    if (OLRXAPI_OK == get_data(argsGetData)):
        Bus1Hnd = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_nBus2Hnd
    if (OLRXAPI_OK == get_data(argsGetData)):
        Bus2Hnd = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dR
    if (OLRXAPI_OK == get_data(argsGetData)):
        dR = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dX
    if (OLRXAPI_OK == get_data(argsGetData)):
        dX = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dR0
    if (OLRXAPI_OK == get_data(argsGetData)):
        dR0 = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dX0
    if (OLRXAPI_OK == get_data(argsGetData)):
        dX0 = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dLength
    if (OLRXAPI_OK == get_data(argsGetData)):
        dLength = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_sName
    if (OLRXAPI_OK == get_data(argsGetData)):
        sName = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["hnd"] = Bus1Hnd
    argsGetData["token"] = BUS_dKVnominal
    if (OLRXAPI_OK == get_data(argsGetData)):
        dKV = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())

    BusHndList = (c_int*100)(0)

    BusListCount = 2
    if Bus1Hnd > Bus2Hnd:
        BusHndList[0] = Bus2Hnd
        BusHndList[1] = Bus1Hnd
    else:
        BusHndList[0] = Bus1Hnd
        BusHndList[1] = Bus2Hnd

    aLine1 = OlxAPI.FullBusName(Bus1Hnd) + " - " + OlxAPI.FullBusName(Bus2Hnd) + ": Z=" + printImpedance(dR,dX,dKV) + " Zo=" + printImpedance(dR0,dX0,dKV) + " L=" + str(dLength)
    ProcessedHnd[nLineHnd-hndOffset] = 1

    # find tap segments on Bus1 side
    BusHnd  = Bus1Hnd
    while (True):
        LineHnd = FindTapSegmentAtBus(BusHnd, ProcessedHnd, hndOffset, sName)
        if LineHnd == 0:
            break
        ProcessedHnd[LineHnd-hndOffset] = 1
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_dR
        if (OLRXAPI_OK == get_data(argsGetData)):
            dRn = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dX
        if (OLRXAPI_OK == get_data(argsGetData)):
            dXn = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dR0
        if (OLRXAPI_OK == get_data(argsGetData)):
            dR0n = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dX0
        if (OLRXAPI_OK == get_data(argsGetData)):
            dX0n = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dLength
        if (OLRXAPI_OK == get_data(argsGetData)):
            dL = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_nBus2Hnd
        if (OLRXAPI_OK == get_data(argsGetData)):
            BusFarHnd = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if BusFarHnd == BusHnd:
            argsGetData = {}
            argsGetData["hnd"] = LineHnd
            argsGetData["token"] = LN_nBus1Hnd
            if (OLRXAPI_OK == get_data(argsGetData)):
                BusFarHnd = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
        dLength = dLength + dL
        dR  = dR  + dRn
        dX  = dX  + dXn
        dR0 = dR0 + dR0n
        dX0 = dX0 + dX0n
        aLine = OlxAPI.FullBusName(BusHnd) + " - " + OlxAPI.FullBusName(BusFarHnd) + ": Z=" + printImpedance(dRn,dXn,dKV) + " Zo=" + printImpedance(dR0n,dX0n,dKV) + " L=" + str(dL)
        print("Segment: " + aLine)
        ProcessedHnd[LineHnd-hndOffset] = 1
        BusHndList[BusListCount] = BusHnd
        BusListCount = BusListCount+1
        BusHnd  = BusFarHnd
        BusHndList[BusListCount] = BusFarHnd
        BusListCount = BusListCount+1

    # find tap segments on Bus1 side
    BusHnd  = Bus2Hnd
    while (True):
        LineHnd = FindTapSegmentAtBus(BusHnd, ProcessedHnd, hndOffset, sName)
        if LineHnd == 0:
            break
        ProcessedHnd[LineHnd-hndOffset] = 1
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_dR
        if (OLRXAPI_OK == get_data(argsGetData)):
            dRn = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dX
        if (OLRXAPI_OK == get_data(argsGetData)):
            dXn = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dR0
        if (OLRXAPI_OK == get_data(argsGetData)):
            dR0n = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dX0
        if (OLRXAPI_OK == get_data(argsGetData)):
            dX0n = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dLength
        if (OLRXAPI_OK == get_data(argsGetData)):
            dL = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_nBus2Hnd
        if (OLRXAPI_OK == get_data(argsGetData)):
            BusFarHnd = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if BusFarHnd == BusHnd:
            argsGetData = {}
            argsGetData["hnd"] = LineHnd
            argsGetData["token"] = LN_nBus1Hnd
            if (OLRXAPI_OK == get_data(argsGetData)):
                BusFarHnd = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
        dLength = dLength + dL
        dR  = dR  + dRn
        dX  = dX  + dXn
        dR0 = dR0 + dR0n
        dX0 = dX0 + dX0n
        aLine = OlxAPI.FullBusName(BusHnd) + " - " + OlxAPI.FullBusName(BusFarHnd) + ": Z=" + printImpedance(dRn,dXn,dKV) + " Zo=" + printImpedance(dR0n,dX0n,dKV) + " L=" + str(dL)
        print("Segment: " + aLine)
        ProcessedHnd[LineHnd-hndOffset] = 1
        BusHndList[BusListCount] = BusHnd
        BusListCount = BusListCount+1
        BusHnd  = BusFarHnd
        BusHndList[BusListCount] = BusFarHnd
        BusListCount = BusListCount+1


    if BusListCount > 2:
        print("Segment: " + aLine1)
        Changed = 1
        while (Changed > 0):
            Changed = 0
            for ii in range(0,BusListCount-2):
                if BusHndList[ii] > BusHndList[ii+1]:
                    nTemp = BusHndList[ii]
                    BusHndList[ii] = BusHndList[ii+1]
                    BusHndList[ii+1] = nTemp
                    Changed = 1

        for ii in range(0,BusListCount-2):
            if BusHndList[ii] == BusHndList[ii+1]:
                BusHndList[ii] = 0
                BusHndList[ii+1] = 0

        jj = 0
        nflg = 0
        for ii in range(0,BusListCount-1):
            if BusHndList[ii] > 0:
                BusHndList[jj] = BusHndList[ii]
                jj = jj + 1
            if jj == 2:
                nflg = 1
                break

        if nflg == 1:
            aLine1 = OlxAPI.FullBusName(busHndList[0]) + " - " + OlxAPI.FullBusName(busHndList[1]) + ": Z=" + printImpedance(dR,dX,dKV) + " Zo=" + printImpedance(dR0,dX0,dKV) + " L=" + str(dLength)
    print("Line: " + aLine1)

def FindTapSegmentAtBus( BusHnd, ProcessedHnd, hndOffset, sName ):
    FindTapSegmentAtBus = 0
    argsGetData = {}
    argsGetData["hnd"] = BusHnd
    argsGetData["token"] = BUS_nTapBus
    if (OLRXAPI_OK == get_data(argsGetData)):
        TapCode = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    if TapCode == 0:
        return 0
    BranchHnd = c_int(0)
    while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( BusHnd, c_int(TC_BRANCH), byref(BranchHnd) ) ) :
        argsGetData = {}
        argsGetData["hnd"] = BranchHnd.value
        argsGetData["token"] = BR_nType
        if (OLRXAPI_OK == get_data(argsGetData)):
            TypeCode = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if TypeCode != TC_LINE:
            continue
        argsGetData = {}
        argsGetData["hnd"] = BranchHnd.value
        argsGetData["token"] = BR_nHandle
        if (OLRXAPI_OK == get_data(argsGetData)):
            LineHnd = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if ProcessedHnd[LineHnd - hndOffset] == 1:
            continue
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_sName
        if (OLRXAPI_OK == get_data(argsGetData)):
            sNameThis = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if sNameThis == sName:
            return LineHnd
        if sNameThis[:3] == "[T]":
            continue
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_sID
        if (OLRXAPI_OK == get_data(argsGetData)):
            sIDThis = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if sIDThis == "T":
            continue
        FindTapSegmentAtBus = LineHnd
    return FindTapSegmentAtBus

def printImpedance(dR, dX, dKV):
    dMag = math.sqrt(dR*dR + dX*dX)*dKV*dKV/100.0
    if dR != 0.0:
        dAng = math.atan(dX/dR)*180/3.14159
    else:
        if dX > 0:
            dAng = 90
        else:
            dAng = -90
    aLine = "{0:.5f}".format(dR) + "+j" + "{0:.5f}".format(dX) + "pu(" + "{0:.2f}".format(dMag) + "@" + "{0:.2f}".format(dAng) + "Ohm)"
    return aLine

def GetRemoteTerminals( BranchHnd, TermsHnd ):
    """
    Purpose: Find all remote ends of a line. All taps are ignored. Close switches are included
    Usage:
      BranchHnd [in] branch handle of the local terminal
      TermsHnd  [in] array to hold list of branch handle at remote ends
    Return: Number of remote ends
    """
    ListSize = 0
    MXSIZE =100
    TempListSize = 1
    TempBrList = (c_int*(MXSIZE))(0)
    TempBrList[TempListSize-1] = BranchHnd
    while (TempListSize > 0):
        NearEndHnd = TempBrList[TempListSize-1]
        TempListSize = TempListSize - 1
        ListSize = FindOppositeBranch( NearEndHnd, TermsHnd, TempBrList, TempListSize, MXSIZE )
    return ListSize

def FindOppositeBranch( NearEndBrHnd, OppositeBrList, TempBrList, TempListSize, MXSIZE ):
    TempListSize2 = 0
    ListSize = 0
    TempList2 = (c_int*(MXSIZE))(0)
    argsGetData = {}
    argsGetData["hnd"] = NearEndBrHnd
    argsGetData["token"] = BR_nInService
    if (OLRXAPI_OK == get_data(argsGetData)):
        nStatus = argsGetData["data"]
    if nStatus != 1:
        return 0
    argsGetData = {}
    argsGetData["hnd"] = NearEndBrHnd
    argsGetData["token"] = BR_nBus2Hnd
    if (OLRXAPI_OK == get_data(argsGetData)):
        nBus2Hnd = argsGetData["data"]
    else:
        return 0
    for ii in range(0, TempListSize2):
        if TempListSize2[ii] == nBus2Hnd:
            return 0
    if (TempListSize2-1 == MXSIZE):
        print("Ran out of buffer spase. Edit script code to incread MXSIZE" )
    TempListSize2 = TempListSize2 + 1
    TempList2[TempListSize2-1] = nBus2Hnd

    argsGetData = {}
    argsGetData["hnd"] = NearEndBrHnd
    argsGetData["token"] = BR_nHandle
    if (OLRXAPI_OK == get_data(argsGetData)):
        nThisLineHnd = argsGetData["data"]
    else:
        return 0
    argsGetData = {}
    argsGetData["hnd"] = nBus2Hnd
    argsGetData["token"] = BUS_nTapBus
    if (OLRXAPI_OK == get_data(argsGetData)):
        nTapBus= argsGetData["data"]
    else:
        return 0
    nBranchHnd = c_int(0);
    if (nTapBus != 1):
        while ( OLRXAPI_OK == OlxAPI.GetBusEquipment( nBus2Hnd, c_int(TC_BRANCH), byref(nBranchHnd) ) ) :
            argsGetData = {}
            argsGetData["hnd"] = nBranchHnd.value
            argsGetData["token"] = BR_nHandle
            if (OLRXAPI_OK == get_data(argsGetData)):
                nLineHnd = argsGetData["data"]
            else:
                return 0
            if (nThisLineHnd == nLineHnd):
                break;
        ListSize = ListSize + 1
        OppositeBrList[ListSize-1] = nBranchHnd.value
    else:
        while ( OLRXAPI_OK == OlxAPI.GetBusEquipment( nBus2Hnd, c_int(TC_BRANCH), byref(nBranchHnd) ) ) :
            argsGetData = {}
            argsGetData["hnd"] = nBranchHnd.value
            argsGetData["token"] = BR_nHandle
            if (OLRXAPI_OK == get_data(argsGetData)):
                nLineHnd = argsGetData["data"]
            else:
                return 0
            if (nThisLineHnd == nLineHnd):
                continue;
            argsGetData["token"] = BR_nType
            if (OLRXAPI_OK == get_data(argsGetData)):
                nType = argsGetData["data"]
            else:
                return 0
            if ((nType != TC_LINE) and (nType != TC_SWITCH)):
                continue
            if (nType == TC_SWITCH):
                argsGetData = {}
                argsGetData["hnd"] = nLindHnd
                argsGetData["token"] = SW_nInService
                if (OLRXAPI_OK == get_data(argsGetData)):
                    nStatus = argsGetData["data"]
                else:
                    return 0
                if (nStatus != 1):
                    continue
                argsGetData["token"] = SW_nStatus
                if (OLRXAPI_OK == get_data(argsGetData)):
                    nStatus = argsGetData["data"]
                else:
                    return 0
                if (nStatus != 1):
                    continue
            if (nType == TC_LINE):
                argsGetData = {}
                argsGetData["hnd"] = nLineHnd
                argsGetData["token"] = LN_nInService
                if (OLRXAPI_OK == get_data(argsGetData)):
                    nStatus = argsGetData["data"]
                else:
                    return 0
                if (nStatus != 1):
                    continue
            if (TempListSize-1 == MXSIZE):
                print("Ran out of buffer spase. Edit script code to incread MXSIZE" )
                return 0
            TempListSize = TempListSize + 1
            TempBrList[TempListSize] = nBranchHnd
    return ListSize

def getBusByBranch(hndBr):
    """
    Find bus of branch
        Args :
            hndBr :  branch handle

        Returns:
            bus[nBus1Hnd,nBus2Hnd, (nBus3Hnd)]

        Raises:
            OlrxAPIException
    """
    bres = []
    b1 = getEquipementData([hndBr],[BR_nBus1Hnd],VT_INTEGER)[0][0]
    bres.append(b1)
    #
    b2 = getEquipementData([hndBr],[BR_nBus2Hnd],VT_INTEGER)[0][0]
    bres.append(b2)
    try:
        b3 = getEquipementData([hndBr],[BR_nBus3Hnd],VT_INTEGER)[0][0]
        bres.append(b3)
    except:
        pass
    #
    return bres

def getBusByEquipement(ehnd,TC_Type):
    """
    Find bus of equipement
        Args :
            ehnd :  equipement handle
            TC_Type: TC_LINE | TC_SCAP | TC_PS | TC_SWITCH | TC_XFMR | TC_XFMR3

        Returns:
            bus[nBus1Hnd,nBus2Hnd, (nBus3Hnd)]

        Raises:
            OlrxAPIException
    """
    if TC_Type == TC_LINE:
        busHnd = [LN_nBus1Hnd,LN_nBus2Hnd]
    #
    elif TC_Type == TC_SCAP:
        busHnd = [SC_nBus1Hnd,SC_nBus2Hnd]
    #
    elif TC_Type == TC_PS:
        busHnd = [PS_nBus1Hnd,PS_nBus2Hnd]
    #
    elif TC_Type == TC_SWITCH:
        busHnd = [SW_nBus1Hnd,SW_nBus2Hnd]
    #
    elif TC_Type == TC_XFMR:
        busHnd = [XR_nBus1Hnd,XR_nBus2Hnd]
    #
    elif TC_Type == TC_XFMR3:
        busHnd = [X3_nBus1Hnd,X3_nBus2Hnd,X3_nBus3Hnd]
    else:
        raise OlrxAPIException("Error TC_Type")
    #
    busRes = []
    for b1 in busHnd:
        busRes.append(getEquipementData([ehnd],[b1],VT_INTEGER)[0][0])
    #
    return busRes


def getEquipementData(ehnd,paraCode,VT_type):
    """
    Find data of element [] (line/bus/...)
        Args :
            ehnd []:  handle element
            nParaCode [] = code data (BUS_,LN_,...)
            VT_type = VT_STRING | VT_DOUBLE | VT_INTEGER|
                      VT_ARRAYSTRING | VT_ARRAYDOUBLE | VT_ARRAYINT

        Returns:
            data [len(paraCode)] [len(bhnd)]

        Raises:
            OlrxAPIException
   """
    # type code
    if VT_type ==VT_INTEGER:
        val1 = c_int(0)
    elif VT_type ==VT_DOUBLE:
        val1 = c_double(0)
    elif VT_type ==VT_STRING:
        val1 = create_string_buffer(b'\000' * 128)
    else:
        raise OlxAPI.OlrxAPIException("Error of VT_type")
    # get data
    res = []
    for paraCode1 in paraCode:
        r1 = []
        for ehnd1 in ehnd:
            if ( OLRXAPI_FAILURE == OlxAPI.GetData( ehnd1, c_int(paraCode1), byref(val1) ) ) :
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            r1.append(val1.value)
        #
        res.append(r1)
    #
    return res

def getBusEquipmentData(bhnd,paraCode,VT_type):
    """
    Retrieves the handle of all equipment of a given type [] (paraCode)
    that is attached to bus [].

        Args :
            bhnd :  [bus handle]
            nParaCode [] = code data (BR_nHandle,GE_nBusHnd...)
            VT_type = VT_STRING | VT_DOUBLE | VT_INTEGER|
                      VT_ARRAYSTRING | VT_ARRAYDOUBLE | VT_ARRAYINT

        Returns:
           [][][]  = [len(paraCode)] [len(bhnd)] [len(all equipement)]

        Raises:
            OlrxAPIException
   """
    # type code
    if VT_type ==VT_INTEGER:
        val1 = c_int(0)
    elif VT_type ==VT_DOUBLE:
        val1 = c_double(0)
    elif VT_type ==VT_STRING:
        val1 = create_string_buffer(b'\000' * 128)
    else:
        raise OlxAPI.OlrxAPIException("Error of VT_type")

    # get data
    r0 = []
    for paraCode1 in paraCode:
        r1 = []
        for bhnd1 in bhnd:
            r2 = []
            while ( OLRXAPI_OK == OlxAPI.GetBusEquipment( bhnd1, c_int(paraCode1), byref(val1) ) ) :
                r2.append(val1.value)
            r1.append(r2)
        r0.append(r1)
    return r0

def getEquipementHandle(TC_type):
    """
    Find all handle of element [] (line/bus/...)
        Args :
            TC_type :  type element (TC_LINE, TC_BUS,...)

        Returns:
            all handle

        Raises:
            OlrxAPIException
   """
    res = []
    hndBr = c_int(0)
    while ( OLRXAPI_OK == OlxAPI.GetEquipment(c_int(TC_type), byref(hndBr) )) :
        res.append(hndBr.value)
    return res

def branchIsInType(hndBr,typeConsi):
    """
    test if branch (in service) is in type defined by typeConsi:  [TC_LINE,TC_SWITCH,TC_SCAP, TC_PS, TC_XFMR,TC_XFMR3]
                       Close switches are tested
    """
    da1 =  getEquipementData([hndBr],[BR_nType,BR_nInService],VT_INTEGER)
    type1 = da1[0][0]
    inSer = da1[1][0]
    if inSer ==0:
        return False
    #
    if type1 not in typeConsi:
        return False

    if (TC_SWITCH in typeConsi) and (type1 ==TC_SWITCH):# if Switch
        swHnd = getEquipementData([hndBr],[BR_nHandle],VT_INTEGER)[0][0]
        status =  getEquipementData([swHnd],[SW_nStatus],VT_INTEGER)[0][0]
        if status==0: #OPEN
            return False
    return True

def branchesNextToBranch(hndBr,typeConsi):
    """
    Purpose: Find list branches next to a branch
             All taps are ignored. Close switches are included

        Args :
            hndBr :  branch handle
            typeConsi: type of branche considered
                       [TC_LINE,TC_SWITCH,TC_SCAP, TC_PS, TC_XFMR]

        returns :
            nTap: nTap of Bus2 (of hndBr)
            br_Self: branch inverse (by bus) of hndBr
            br_res []: list branch next
                                              | br_res
                                              |
                                  hndBr       |      br_res
        Illustration:       Bus1------------Bus2------------
                                  br_Self     |
                                              |
                                              | br_res
        Raises:
            OlrxAPIException
    """
    br_Self = -1
    br_res = []
    #
    b2 = getEquipementData([hndBr],[BR_nBus2Hnd],VT_INTEGER)[0][0]
    # get nTap
    nTap = getEquipementData([b2],[BUS_nTapBus],VT_INTEGER)[0][0]
    # equipement handle of hndBr
    equiHnd_0 = getEquipementData([hndBr],[BR_nHandle],VT_INTEGER)[0][0]
    # all branch connect to b2
    allBr0 = getBusEquipmentData([b2],[TC_BRANCH],VT_INTEGER)[0][0]
    #
    allBr = []
    for br1 in allBr0:
        if branchIsInType(br1,typeConsi): #test branch is in type considered, Switch close
            allBr.append(br1)
    #
    equiHnd_all = getEquipementData(allBr,[BR_nHandle],VT_INTEGER)[0]
    #
    for i in range(len(allBr)):
        if equiHnd_all[i]== equiHnd_0: # branch same equipement
            br_Self = allBr[i]
        else:
            br_res.append(allBr[i])
    #
    return nTap,br_Self,br_res

def lineComponents(hndBr,typeConsi): # Without XFMR3
    """
    Purpose: find list branches of line (start by hndBr)
            All taps are ignored. Close switches are included

        Args :
            hndBr :  branch handle (start)
            typeConsi: type considered as component of line
            [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR]

        returns :
            list of branches

        Raises:
            OlrxAPIException
    """
    brA_res = [[hndBr]] # result
    bra = [hndBr]# for each direction

    while True:
        bra_in = []
        for br1 in bra:
            #
            nTap,br_Self,br_res = branchesNextToBranch(br1, typeConsi)
            if nTap ==0 or len(br_res)==0: # finish
                for i in range(len(brA_res)):
                    if br1 in brA_res[i]:
                        brA_res[i].append(br_Self)
            else:
                if len(br_res)==1: # tap 2
                    for i in range(len(brA_res)):
                        if br1 in brA_res[i]:
                            brA_res[i].append(br_res[0])
                            bra_in.append(br_res[0])  # continue this direction
                else: # tap>3
                    k = -1
                    for i in range(len(brA_res)):
                        if br1 in brA_res[i]:
                            k = i
                    #
                    for i in range(1,len(br_res)):
                        bri = []
                        bri.extend(brA_res[k])
                        bri.append(br_res[i])
                        brA_res.append(bri)
                    #
                    brA_res[k].append(br_res[0])
                    bra_in.extend(br_res)
        # finish or not
        if len(bra_in)==0:
            break
        # continue
        bra = []
        bra.extend(bra_in)
    #
    return brA_res

def getRemoteTerminals(hndBr):
    """
    Purpose: Find all remote end of a branch
             All taps are ignored.
             Close switches are included
             Out of service branches are ignored

        Args :
            hndBr :  branch handle

        returns :
            bus_res [] list of terminal bus

        Raises:
            OlrxAPIException
    """
    inSer1 = getEquipementData([hndBr],[BR_nInService],VT_INTEGER)[0][0]
    if inSer1 == 2:
        return []

    bus_res = set()
    setEqui = set()
    equiHnd1 = getEquipementData([hndBr],[BR_nHandle],VT_INTEGER)[0][0]
    setEqui.add(equiHnd1)
    #
    b123 = getBusByBranch(hndBr)
    b23 = b123[1:]
    #
    while True:
        bn23_in = []
        for b2 in b23:
            nTap = getEquipementData([b2],[BUS_nTapBus],VT_INTEGER)[0][0]
            allBr = getBusEquipmentData([b2],[TC_BRANCH],VT_INTEGER)[0][0]
            if nTap==0 or len(allBr)==1:
                bus_res.add(b2)
            else:
                for br1 in allBr:
                    equi1 = getEquipementData([br1],[BR_nHandle],VT_INTEGER)[0][0]
                    if (equi1 not in setEqui):
                        inSer1 = getEquipementData([br1],[BR_nInService],VT_INTEGER)[0][0]
                        type1 =  getEquipementData([br1],[BR_nType],VT_INTEGER)[0][0]
                        statusSW = 1
                        if type1 == TC_SWITCH: # if Switch
                            swHnd = getEquipementData([br1],[BR_nHandle],VT_INTEGER)[0][0]
                            statusSW =  getEquipementData([swHnd],[SW_nStatus],VT_INTEGER)[0][0]
                        #
                        if statusSW==0 and len(allBr)==2: # finish
                            bus_res.add(b2)
                        elif statusSW>0 and inSer1==1 :
                            setEqui.add(equi1)
                            b123 = getBusByBranch(br1)
                            bn23_in.extend(b123[1:])
        #update b23
        b23 = []
        b23.extend(bn23_in)
        if len(b23)==0:
            break
    return list(bus_res)

#
def saveString2File(nameFile,sres):
    try:
        text_file = open(nameFile, "w")
        text_file.write(sres)
        text_file.close()
    except:
        pass

def read_File_text (sfile):
    # read file text to array
    ar = []
    try:
        ins = open( sfile, "r" )
        for line in ins:
            ar.append(line.replace("\n",""))
        ins.close()
    except OSError:
        raise Exception(OSError.strerror)
    #
    return ar

def saveAsOlr(fileNew):
    fileNew = fileNew.replace("/","\\")
    if fileNew:
        if OLRXAPI_FAILURE == OlxAPI.SaveDataFile(fileNew):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        return True
    return False

def compare2FileText(file1,file2):
    # compare 2 file Text:
    # returns: arrays of number of line different
    ar1 = read_File_text(file1)
    ar2 = read_File_text(file2)
    dif = []
    for i in range(max(len(ar1),len(ar2))):
        a1 = ""
        a2 = ""
        try:
            a1 = ar1[i]
        except:
            pass
        try:
            a2 = ar2[i]
        except:
            pass
        #
        a1 = a1.strip()
        a2 = a2.strip()
        if a1!=a2:
            dif.append(i+1)
    return dif

def unit_test_compare(PATH_FILE,PY_FILE,sres):
    """
    tool for unit test
    - save result to file.txt
    - compare result with REF
    """
    fileRes = os.path.join(PATH_FILE,PY_FILE.replace(".py",".txt"))
    #
    saveString2File(fileRes,sres)
    # compare.
    fileREF = fileRes.replace(".txt","_REF.txt")
    #
    dif = compare2FileText(fileREF,fileRes) # return array of line number with difference
    if len(dif)==0:
        print("PASS unit test: "+ PY_FILE, "(", os.path.basename(fileRes), "==",os.path.basename(fileREF),")")
    else:
        print("PROBLEM unit test: " + PY_FILE, "(", os.path.basename(fileRes), "!=",os.path.basename(fileREF),")")
        print("\tdifferences in lines: ",dif)


def setData(hnd,paraCode,value,VT_type):
    """
     Set data for equipement (bus,line,...)
        Args :
            hnd:  handle element
            nParaCode = code data (BUS_,LN_,...)
            value: value to be set
            VT_type = VT_STRING | VT_DOUBLE | VT_INTEGER|
                      VT_ARRAYSTRING | VT_ARRAYDOUBLE | VT_ARRAYINT

        Returns:
            None

        Raises:
            OlrxAPIException
    """
    if VT_type ==VT_INTEGER:
        val1 = c_int(value)
    elif VT_type ==VT_DOUBLE:
        val1 = c_double(value)
    elif VT_type ==VT_STRING:
        val1 = c_char_p(value)
    else:
        raise OlxAPI.OlrxAPIException("Error of VT_type")
    #
    if OLRXAPI_OK != OlxAPI.SetData( c_int(hnd), c_int(paraCode), byref(val1) ):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())

def postData(hnd):
    """Post network object data
    """
    # Validation
    if OLRXAPI_OK !=  OlxAPI.PostData(c_int(hnd)):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    return OLRXAPI_OK


class BusSearch:
    def __init__(self,gui):
        self.gui = gui
        self.bhnd = []
        self.s1 = ""
        self.s2 = ""
        self.exact = False
    #
    def getData(self):
        if len(self.bhnd)>0:
            return
        self.updateData()

    def updateData(self):
        self.bhnd = getEquipementHandle(TC_BUS)
        self.kv   = getEquipementData(self.bhnd,[BUS_dKVnominal],VT_DOUBLE)[0]
        self.name = getEquipementData(self.bhnd,[BUS_sName]     ,VT_STRING)[0]
        for i in range(len(self.name)):
            self.name[i]= self.name[i].upper()
        #
        self.num  = getEquipementData(self.bhnd,[BUS_nNumber]   ,VT_INTEGER)[0]
    #
    def searchBy_BusNumber(self,bNum):
        self.s1 = "Bus search by number"
        self.s2 = "number = "+ str(bNum)
        #
        bhnd = OlxAPI.FindBusNo(bNum)
        if bhnd>0:
            self.exact = True
            return self.__selectEnd([bhnd])
        #

        self.getData()
        num = [bNum]
        num.extend(self.num)
        num.sort()
        i = num.index(bNum)
        if i==0:
            return self.__selectEnd2([num[1]])
        if i==len(num)-1:
            return self.__selectEnd2([num[i-1]])
        #
        return self.__selectEnd2([num[i-1],num[i+1]])

    def searchBy_NameKv(self,bname,kV):
        self.s1 = "Bus search by (name, kV)"
        self.s2 = "name , kV = " + bname + " , " + str(kV)
        #
        bhnd  = OlxAPI.FindBus(bname,kV)
        if bhnd>0:
            self.exact = True
            return self.__selectEnd([bhnd])
        #
        self.getData()
        self.na = [None]* len(self.bhnd)
        #
        bnameUp = bname.upper()
        #
        bn1 =  bnameUp[:1]
        kr = self.__searchBy_NameKv_1(bn1,kV,deltaKv = 1)
        if len(kr)== 0: # extend deltaKv = 3
            kr = self.__searchBy_NameKv_1(bn1,kV,deltaKv = 3)
        if len(kr)== 0: # extend deltaKv = 5
            kr = self.__searchBy_NameKv_1(bn1,kV,deltaKv = 10)
        #
        if len(kr)== 0:
            return self.__selectEnd([0])
        #
        if len(kr)==1 or len(bname)==1: # finish search
            return self.__selectEnd1(kr)

        for j in range(1,len(bname)):
            # search 2 start with 2s
            bn1 = bnameUp[j:j+1]
            kr2 = self.__searchBy_NameKv_2(bn1,kr)
            #
            if len(kr2)==0: # start with 1s, 1s in
                kr2 = self.__searchBy_NameKv_3(bn1,kr)
            if len(kr2)==0:
                return self.__selectEnd1(kr)
            #
            if len(kr2)==1 or len(bname)==(j+1): # finish search
                return self.__selectEnd1(kr2)
            kr = []
            kr.extend(kr2)
        #
        return self.__selectEnd1(kr)

    def __searchBy_NameKv_1(self,bn1,kV,deltaKv):
        kr1 = []
        for i in range(len(self.name)):
            if (self.name[i]).startswith(bn1):
                if abs(kV-self.kv[i])<deltaKv:
                    kr1.append(i)
                    self.na[i]= self.name[i][1:]
        return kr1
    #
    def __searchBy_NameKv_2(self,bn1,kr1):
        kr2 = []
        for i in kr1:
            if (self.na[i]).startswith(bn1):
                kr2.append(i)
                self.na[i]= self.na[i][1:]
        return kr2
    #
    def __searchBy_NameKv_3(self,bn1,kr1):
        kr2 = []
        for i in kr1:
            s= self.na[i]
            if bn1 in self.na[i]:
                kr2.append(i)
        return kr2
    #
    def __selectEnd1(self,kr):
        ba = []
        for k in kr:
            ba.append(self.bhnd[k])
        return self.__selectEnd(ba)
    #
    def __selectEnd2(self,num):
        ba = []
        for n1 in num:
            bhnd = OlxAPI.FindBusNo(n1)
            ba.append(bhnd)
        return self.__selectEnd(ba)
    #
    def __selectEnd(self,ba):
        print(self.s1)
        print("\t"+self.s2)
        #
        if self.gui==1:
            root = tk.Tk()
            try:
                root.wm_iconbitmap(OLXAPI_DLL_PATH+"ASPEN.ico")
            except:
                pass
        #
        if len(ba)==1:
            if ba[0]==0:
                s1 = "No bus found"
            else:
                if self.exact:
                    s1 = "Bus found: " + OlxAPI.FullBusName(ba[0])
                else:
                    s1 = "Bus found (nearest): " + OlxAPI.FullBusName(ba[0])
            #
            print("\t"+s1)
            #
            if self.gui ==1:
                root.withdraw()
                #
                tkm.showinfo(self.s1,self.s2+"\n\n"+s1)
                root.destroy()
            return ba[0]
        # multipe result
        data = []
        for bi in ba:
            data.append (OlxAPI.FullBusName(bi))
        #
        s1 = "Bus found (nearest): "
        for d1 in data:
            s1 += "\n\t\t"+ d1
        print("\t"+s1)
        if self.gui==0:
            print("\tSelected: ",OlxAPI.FullBusName(ba[0]))
            return ba[0]
        #
        try:
            root.wm_title(self.s1)
            ws = root.winfo_screenwidth()
            hs = root.winfo_screenheight()
            root.geometry("300x200+"+ str(int((ws-300)/2))+"+" + str(int((hs-200)/2)))
            w1 = tk.Label(root, text=self.s2)
            w1.place(x=50, y=10)
            #
            w2 = tk.Label(root, text="Select the nearest bus")
            w2.place(x=50, y=40)
            #
            select = []
            d = self.APP(root,data,select)
            root.mainloop()
            i = data.index(select[0])
            return ba[i]
        except:
            print("\tSelected: ",OlxAPI.FullBusName(ba[0]))
            return ba[0]
    #
    class APP(tk.Frame):
        def __init__(self, parent,data,select):
            tk.Frame.__init__(self, parent)
            self.parent = parent
            self.data = data
            self.select = select
            self.initGUI()
        #
        def initGUI(self):
            #
            self.cb = ttk.Combobox(self.parent,width=25, values=self.data)
            self.cb.place(x=50, y=70)
            self.cb.bind("<<ComboboxSelected>>",self.co3)
            self.cb.current(0)
            #
            button_OK = tk.Button(self.parent,text =   '     OK     ',command=self.run_OK)
            button_OK.place(x=120, y=150)
        #
        def co3(self,event):
            x = self.cb.get()
            print("\tSelect:",x)
        #
        def run_OK(self):
            x = self.cb.get()
            self.select.append(x)
            self.parent.destroy()

