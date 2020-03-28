"""
These are some rough ideas for the trial project that we discussed in the web meeting.
Feel free to make changes and propose your own

1. OLR network visualizer:
    develop a Python app with graphical UI to display
    all transmission lines and/or transformers that are located at the near end and the far end of a selected line.
    Bonus: correctly handle parallel lines and lines with tap buses

2. Network tracer:
    given a low voltage bus (such a 15kV),
    print a list of all connected lines and transformers in the network with the kV up to next 1, 2, 3 level (such as 35, 69, 110)

3. Network fault comparison tool:
    given a OLR file and fault solution report in TXT or CSV format,
    repeat the simulation to produce a new solution report and generate a side-by-side comparison of two results

----------------
Phuong'solution:
    2. networkTracer(bsName,bsKV,nLevel)
        given a bus (bsName,bsKV)
        print a list of all branches/bus in the network connected to the given bus with nLevel
            nLevel = 1: all branches connected directly to the given bus
##                        all bus in another side of branch Level 1
            nLevel = 2: all branches connected to bus Level 1 (except branches level-1)
                        all bus in another side of branch Level 2
            .......
            => to improve: nLevel defined by station

    3. faultComparisonTool(olrFilePath,csvResFile,bsName,bsKV,nLevel,sCSV,nRound
        given:  - olrFilePath: OLR file
                - csvResFile: CSV file
                - a bus (bsName,bsKV): fault at this bus
                - nLevel: print result fault by nLevel connected to fault bus => networkTracer
                - sCSV: delimited of CSV file ","
                - nRound: round result example u = 12.23 (nRound= 2)
        results:
                - if the csvResFile does not exist => create a new CSV file
                - if csvResFile is a file: add the new result to compare
                - for the moment: just compare the bus voltage
                -=> to improve: compare branch current
"""
from __future__ import print_function
from pyOlrxAPILib import *
import time
import string

#  -----------------------------------------------------------------------------------------------------------------------
def read_File_text (sfile):
    array = []
    try:
        ins = open( sfile, "r" )
        for line in ins:
            array.append(line.replace("\n",""))
        ins.close()
    #
    except:
        raise("File not found :" + sfile +"'")
    #
    return array
#  -----------------------------------------------------------------------------------------------------------------------
def write_File_text (sfile , array):
    for i in range(len(array)):
        array[i]+="\n"
    #
    try:
        f = open(sfile, "w")
        try:
            f.writelines(array) # Write a sequence of strings to a file
        except:
            print("\n\nERROR in _write_File_text"+sfile)
        finally:
            f.close()
            return sfile
    except :
        t1 = string.rindex(sfile, ".")
        sn = sfile[0:t1]+"_1."+sfile[t1+1:len(sfile)]
        write_File_text(sn , array)

# Open in excel to visu (only)
def openCSVExcel(csvFile):
    try:
        #kill excel
        os.system("taskkill /im EXCEL.exe")
        os.system('taskkill /f /im EXCEL.exe')
        # excell tool
        import win32com.client
        xlApp = win32com.client.Dispatch("Excel.Application")
        xlApp.Workbooks.Open(csvFile)
        xlApp.Visible = True
    except:
        pass
#
def loadOlrxdllFile():
    #
    dllPath = "C:\\Program Files (x86)\\ASPEN\\1LPFv15\\"
    dllFile = "olxapi.dll"
    try:
        InitOlrxAPI_1(dllPath,dllFile)
        print ("\nASPEN OlrxAPI path: ", dllPath, "\nVersion:", OlrxAPIVersion(), " Build: ", OlrxAPIBuildNumber(), "\n")
    except OlrxAPIException as err:
        raise OlrxAPIException("OlrxAPI Init Error: {0}".format(err))

#----------------------------
def findConnectedBus(bhndAll):
    equiHnd = set() # all equipement connected to bus (bhndAll)
    busHnd  = set() # all bus next to bus (bhndAll)
    #
    for b1 in bhndAll:
        hndBr = c_int(0)
        hndBus2 = c_int(0)
        hndBus3 = c_int(0)
        br_hnd = c_int(0)
        #
        while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( b1, c_int(TC_BRANCH), byref(hndBr) ) ) :
            # get Br_hnd of
            if ( OLRXAPI_FAILURE == OlrxAPIGetData( hndBr, c_int(BR_nHandle), byref(br_hnd))):
                raise OlrxAPIException(OlrxAPIErrorString())
            # bus 2
            if ( OLRXAPI_FAILURE == OlrxAPIGetData( hndBr, c_int(BR_nBus2Hnd), byref(hndBus2) ) ) :
                raise OlrxAPIException(OlrxAPIErrorString())
            #
            equiHnd.add(br_hnd.value)
            busHnd.add(hndBus2.value)
            # bus 3
            if ( OLRXAPI_OK == OlrxAPIGetData( hndBr, c_int(BR_nBus3Hnd), byref(hndBus3) ) ) :
                busHnd.add(hndBus3.value)
    #
    return equiHnd,busHnd

#
def runTracerBus(hnd,nLevel):
    resEqui  = [set()]*(nLevel)
    resBus   = [set()]*(nLevel)
    resEqui[0],resBus[0] = findConnectedBus([hnd])
    #
    for i in range(1,nLevel):
        resEqui[i],resBus[i] = findConnectedBus(resBus[i-1])
        #
        for j in range(i):#update
            resBus[i]  = resBus[i] - resBus[j]
            resEqui[i] = resEqui[i]- resEqui[j]
    #
    return resEqui,resBus

#
def networkTracer(bsName,bsKV,nLevel):
    print("\nRUN networkTracer ---------------------------------------------")
    #
    bhnd   = OlrxAPIFindBus(bsName, bsKV)
    #
    if bhnd == OLRXAPI_FAILURE:
        raise OlrxAPIException(("\nBus "+ bsName+str(bsKV)+ "kV not found"))
        return
    # run core
    resEqui,resBus = runTracerBus(bhnd,nLevel)

    # PRINT RESULT
    print("Ini Bus: ",OlrxAPIFullBusName(bhnd))
    for i in range(nLevel):
        #
        print("Level-",i+1,"----------------------------")
        #
        print("\tBUS -----------------------------")
        for k1 in resBus[i]:
            print ("\t\t",OlrxAPIFullBusName(k1))

        print("\tEQUIPEMENT-----------------------")
        for k1 in resEqui[i]:
            print ("\t",OlrxAPIFullBranchName(k1))
    #

# run bus fault
def run_busFault(busHnd):
    """Report fault on a bus
    """
    hnd = c_int(busHnd)
    fltConn = (c_int*4)(1,0,0,0)   # 3LG, 2LG, 1LG, LL
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
    if OLRXAPI_FAILURE == OlrxAPIDoFault(hnd, fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev):
        raise OlrxAPIException(OlrxAPIErrorString())
    # select first scene
    if OLRXAPI_FAILURE == OlrxAPIPickFault(c_int(SF_FIRST),c_int(9)):
        raise OlrxAPIException(OlrxAPIErrorString())
    #
    print (OlrxAPIFaultDescription(0))

"""get result bus fault
return:
- all voltage at bus - connect to bus Fault with nLevel
- all current of branch- connect to bus Fault with nLevel
"""
def getResult_busFault(bhnd,nLevel):
    # ini result
    volBusMag = {}
    volBusAng = {}
    iBrMag = {}
    iBrAng = {}
    # run tracer
    resEqui,resBus = runTracerBus(bhnd,nLevel)
    #
    vd12Mag = (c_double*12)(0.0)
    vd12Ang = (c_double*12)(0.0)
    vd9Mag  = (c_double*9)(0.0)
    vd9Ang  = (c_double*9)(0.0)
    # voltage bus
    for i in range(len(resBus)):
        for bh in resBus[i]:
            if ( OLRXAPI_FAILURE == OlrxAPIGetSCVoltage( bh, vd9Mag, vd9Ang, c_int(2) ) ) :
                raise OlrxAPIException(OlrxAPIErrorString())
            volBusMag[bh] = vd9Mag[0:3]
            volBusAng[bh] = vd9Ang[0:3]

    # current branch
    for i in range(len(resEqui)):
        for brh in resEqui[i]:
            if ( OLRXAPI_FAILURE == OlrxAPIGetSCCurrent( brh, vd12Mag, vd12Ang, 4 ) ) :
                raise OlrxAPIException(OlrxAPIErrorString())
            iBrMag[brh] = vd12Mag[0:12]
            iBrAng[brh] = vd12Ang[0:12]
    #
    return volBusMag,volBusAng,iBrMag,iBrAng

#
def print2CSV(olrFilePath,csvResFile,volBusMag,volBusAng,iBrMag,iBrAng,nRound,sCSV):
    # keybus
    busKey = {}
    for bh in volBusMag:
        sx = create_string_buffer("\0",128)
        if OLRXAPI_FAILURE == OlrxAPIGetData( bh, c_int(BUS_sName), byref(sx)):
            raise OlrxAPIException(OlrxAPIErrorString())
        # kV
        kVNom = c_double(0.0)
        if OLRXAPI_FAILURE == OlrxAPIGetData( bh, c_int(BUS_dKVnominal), byref(kVNom)):
            raise OlrxAPIException(OlrxAPIErrorString())
        #
        key = sx.value + " " + str(round(kVNom.value,2)) + " kV"
        #
        busKey[bh]  = key

    # Title --------------------------------------------------------------------
    arTitle = []
    arTitle.append("")
    arTitle.append(sCSV+sCSV+"user: "+os.environ.get("USERNAME")+sCSV+sCSV)
    arTitle.append(sCSV+sCSV+"date: "+time.ctime()+sCSV+sCSV)
    arTitle.append(sCSV+sCSV+"olrFile: "+os.path.basename(olrFilePath)+sCSV+sCSV)
    arTitle.append(sCSV+sCSV+ OlrxAPIFaultDescription(0).replace(" ","")+sCSV+sCSV)
##    arTitle.append("BUS_"+sCSV+sCSV+"VA"+sCSV+"VB (VP)"+sCSV+"VC")

    #write in csv result (1st run fault)----------------------------------------
    if not os.path.isfile(csvResFile):
        ar = []
        ar.append("Fault comparison tool")
        for i in range(1,len(arTitle)):
            ar.append(arTitle[i])
        ar.append("BUS_"+sCSV+sCSV+"VA"+sCSV+"VB (VP)"+sCSV+"VC")
        #BUS
        for bh in volBusMag:
            s1 = "BUS_" + sCSV + busKey[bh]
            for i in range(len(volBusMag[bh])):
                s1+= sCSV + str(round(volBusMag[bh][i],nRound))+"@"+str(round(volBusAng[bh][i]))
            ar.append(s1)
    #current BRANCH : TO DO ----------------------------------------------------
        write_File_text(csvResFile,ar)

    # file exist => add to compare results--------------------------------------
    else:
        ars = read_File_text(csvResFile)
        #
        nmax = 0
        #get size Bus_result
        for j in range(len(ars)):
            if str(ars[j]).startswith("BUS_"):
                nmax = str(ars[j]).count(sCSV)
                break
        #
        for i in range(1,len(arTitle)):
            ars[i] += arTitle[i]
        #
        ars[len(arTitle)] += sCSV+sCSV+"VA"+sCSV+"VB (VP)"+sCSV+"VC"
        #
        abusAj = []
        for bh in volBusMag:
            name1 = busKey[bh]
            test = True
            for j in range(len(arTitle),len(ars)):
                if str(ars[j]).startswith("BUS_"+sCSV+name1):
                    test = False
                    nm1 = str(ars[j]).count(sCSV)
                    for i in range(nm1,nmax+1):
                        ars[j]+= sCSV
                    #
                    for i in range(len(volBusMag[bh])):
                        ars[j]+= sCSV + str(round(volBusMag[bh][i],nRound))+"@"+str(round(volBusAng[bh][i]))
            # bus non result in csv
            if test:
                s1 = "BUS_" + sCSV + busKey[bh]
                for i in range(nmax):
                    s1+= sCSV
                #
                for i in range(len(volBusMag[bh])):
                    s1+= sCSV + str(round(volBusMag[bh][i],nRound))+"@"+str(round(volBusAng[bh][i]))
                #
                abusAj.append(s1)
        #
        na = len(ars)
        for j in range(len(ars)):
            if str(ars[len(ars)-j-1]).startswith("BUS_"+sCSV):
                na = len(ars)-j-1
                break
        # add aj bus
        ars2 = []
        ars2.extend(ars[0:na+1])
        ars2.extend(abusAj)
        ars2.extend(ars[na+1:])
        #
        #current BRANCH : TO DO ------------------------------------------------
        #
        write_File_text(csvResFile,ars2)
    #
    return

#
def faultComparisonTool(olrFilePath,csvResFile,bsName,bsKV,nLevel,sCSV,nRound):
    print("\nRUN faultComparisonTool -------------------------------------------")
    # get bus handle
    bhnd   = OlrxAPIFindBus(bsName, bsKV)
    #
    if bhnd == OLRXAPI_FAILURE:
        raise OlrxAPIException(("\nBus "+ bsName + str(bsKV) + "kV not found"))
        return
    # run Fault
    run_busFault(bhnd)

    # get results
    volBusMag,volBusAng,iBrMag,iBrAng = getResult_busFault(bhnd,nLevel)

    # print 2 CSV
    print2CSV(olrFilePath,csvResFile,volBusMag,volBusAng,iBrMag,iBrAng,nRound,sCSV)
    # open il excel
    openCSVExcel(csvResFile)
    return

#
def main():
    olrFilePath = os.getcwd() + "\\SAMPLE30.OLR" # "\\  TUAN09_A0    SAMPLE30
     # load orldll
    loadOlrxdllFile()
    # open olr file
    if OLRXAPI_FAILURE == OlrxAPILoadDataFile(olrFilePath,1):
        raise OlrxAPIException(OlrxAPIErrorString())
    print ("File opened successfully: ", olrFilePath)

    # INPUT
    csvResFile = os.getcwd() + "\\res.csv"
    bsName = "REUSENS"
    bsName = "NEVADA"
##    bsName = "OHIO"

    bsKV   = 132.0      # 132.0
    sCSV = ","
    nRound = 1

    # miniProject 2
    nLevel = 2
    networkTracer(bsName,bsKV,nLevel)

    # miniProject 3
    nLevel = 1
    faultComparisonTool(olrFilePath,csvResFile,bsName,bsKV,nLevel,sCSV,nRound)

#
if __name__ == '__main__':
    t0 = time.time()
    print(time.ctime(t0))
    main()
    print("\nrunTime= " +str(round(time.time()-t0,1))+"[s]")