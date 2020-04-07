"""
Demo usage of ASPEN OlrxAPI in Python.

Purpose:
    Report tielines by Area/Zone

"""
from __future__ import print_function
from AspenPy import*

__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2018, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.1.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

import Tkinter as tk
import tkFileDialog,tkMessageBox,ttk
import sys,os

# GUI
class APP(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        # get all area/zone
        busHnd = OlxAPILib.getEquipementHandle(TC_BUS)
        aza = OlxAPILib.getEquipementData(busHnd,[BUS_nArea,BUS_nZone],VT_INTEGER)
        #
        self.area = list(set(aza[0]))
        self.area.append(0)
        self.area.sort()
        #
        self.zone = list(set(aza[1]))
        self.zone.append(0)
        self.zone.sort()
        #
        self.initGUI()
    #
    def initGUI(self):
        olrFile = os.path.basename(sys.argv[1])
        self.parent.wm_title("Tieline Report: "+olrFile)
        try:
            self.parent.wm_iconbitmap(OLXAPI_DLL_PATH+"aspen.ico")
        except:
            pass
        #
        la1 = tk.Label(self.parent, text="Type")
        la1.pack()
        la1.place(x=30, y=20)
        #
        self.var1 = tk.IntVar()
        self.r1=tk.Radiobutton(self.parent, text="TieLines"   ,variable = self.var1,value=1,command=self.getValLn)
        self.r1.pack(anchor=tk.W)
        self.r2=tk.Radiobutton(self.parent, text="TieBranches",variable = self.var1,value=2,command=self.getValBr)
        self.r2.pack(anchor=tk.W)
        self.r1.place(x=100,y=20)
        self.r2.place(x=200,y=20)
        self.r1.select()
        self.az="Area"
        #
        la2 = tk.Label(self.parent, text="Tie by ")
        la2.pack()
        la2.place(x=30, y=60)
        #
        self.var2 = tk.IntVar()
        self.r3=tk.Radiobutton(self.parent, text="Area",variable = self.var2, value=1,command=self.getValArea)
        self.r3.pack(anchor=tk.W)
        self.r4=tk.Radiobutton(self.parent, text="Zone",variable = self.var2, value=2,command=self.getValZone)
        self.r4.pack(anchor=tk.W)
        self.r3.place(x=100,y=60)
        self.r4.place(x=200,y=60)
        self.r3.select()
        self.LnBr="Line"
        #
        la2 = tk.Label(self.parent, text="From (0=all)")
        la2.pack()
        la2.place(x=30, y=100)
        #
        self.updateCombo()
        #
        button_OK = tk.Button(self.parent,text =   '     OK     ',command=self.run_OK)
        button_OK.place(x=100, y=150)
        #
        button_exit = tk.Button(self.parent,text = '   Cancel   ',command=self.cancel)
        button_exit.place(x=200, y=150)

    #
    def cancel(self):
        # close OLR file @ TODO
        # quit
        self.parent.destroy()
    #
    def updateCombo(self):
        if self.az=="Area":
            self.cb=ttk.Combobox(self.parent, values=self.area,width = 6)
        elif self.az=="Zone":
            self.cb=ttk.Combobox(self.parent, values=self.zone,width = 6)
        #
        self.cb.place(x=100, y=100)
        self.cb.bind("<<ComboboxSelected>>")
        self.cb.current(0)
    #
    def getValArea(self):
        self.az = "Area"
        self.updateCombo()
    #
    def getValZone(self):
        self.az = "Zone"
        self.updateCombo()
    #
    def getValLn(self):
        self.LnBr = "Line"
    #
    def getValBr(self):
        self.LnBr = "Branch"
       #
    def run_OK(self):
        azNum= int(self.cb.get())
        if self.az=="Area":
            nParamCode = BUS_nArea
        elif self.az=="Zone":
            nParamCode = BUS_nZone
        #
        if self.LnBr=="Line":
            res = tielines(nParamCode,azNum)
        elif self.LnBr=="Branch":
            res = tiebranches(nParamCode,azNum)
        #
        print("\n"+res)
# CORE
def tielines(nParamCode,azNum):
    """
     Report tie Lines by area or zone
        Args :
            nParamCode = BUS_nArea tie lines by AREA
                       = BUS_nZone tie lines by ZONE
            azNum : area/zone number (if <=0 all area/zone)

        Returns:
            string result

        Raises:
            OlrxAPIException
    """
    res  = "List of tie lines in the system"
    if nParamCode == BUS_nArea:
        if azNum<=0:
            res  += " (all area) "
        else:
            res  += " (from area = "+str(azNum) +")"
        res += "\nArea1   ,Area2   ,Lines"
    elif nParamCode == BUS_nZone:
        if azNum<=0:
            res  += " (all zone) "
        else:
            res  += " (from zone = "+str(azNum) +")"
        res += "\nZone1   ,Zone2   ,Lines"
    else:
        raise OlxAPI.OlrxAPIException("Error nParamCode = BUS_nArea | BUS_nZone")
    #
    bus = []
    br = []
    hndBr = c_int(0)
    while ( OLRXAPI_OK == OlxAPI.GetEquipment(c_int(TC_LINE), byref(hndBr) )) :
        bus.append(OlxAPILib.getBusByBranch(hndBr,TC_LINE))
        br.append(hndBr.value)

    # get result
    abr = {}
    for k in range(len(br)):
        b1  = bus[k]
        br1 = br[k]
        area = OlxAPILib.getEquipementData(b1,[nParamCode],VT_INTEGER)[0]
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
        res += "\n"+str(a1).ljust(8)+","+str(a2).ljust(8)+","+OlxAPI.FullBranchName(br1)
    return res

#
def tiebranches(nParamCode,azNum):
    """
     Report tie branches by area or zone
        Args :
            nParamCode = BUS_nArea tie branches by AREA
                       = BUS_nZone tie branches by ZONE
            azNum : area/zone number (if <=0 all area/zone)

        Returns:
            string result

        Raises:
            OlrxAPIException
    """
    res  = "List of tie branches in the system"
    if nParamCode == BUS_nArea:
        if azNum<=0:
            res  += " (all area) "
        else:
            res  += " (from area = "+str(azNum) +")"
        res += "\nArea1   ,Area2   ,Branches"
    elif nParamCode == BUS_nZone:
        if azNum<=0:
            res  += " (all zone) "
        else:
            res  += " (from zone = "+str(azNum) +")"
        res += "\nZone1   ,Zone2   ,Branches"
    else:
        raise OlxAPI.OlrxAPIException("Error nParamCode = BUS_nArea | BUS_nZone")

    #
    br,bus = OlxAPILib.getAllBranchesAndBus()
    # get result
    abr = {}
    for k in range(len(br)):
        b1 = bus[k]
        br1 = br[k]
        area = OlxAPILib.getEquipementData(b1,[nParamCode],VT_INTEGER)[0]
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
        res += "\n"+str(a1).ljust(8)+","+str(a2).ljust(8)+","+OlxAPI.FullBranchName(br1)
    # print
    return res


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
        setPara_sys_argv(idx=1,val=OLR_DEFAULT_PATH)
    olrFile = sys.argv[1]
    #
    load_olxapi_dll()
    open_olrFile(readonly=0)
    #
    root = tk.Tk()
    root.geometry("350x200+300+200")
    d = APP(root)
    root.mainloop()




