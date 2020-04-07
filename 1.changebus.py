"""
Demo usage of ASPEN OlrxAPI in Python.

Purpose: Change all bus data: Area/Zone in OLR file
         from one value to another

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
        #
        busHnd = OlxAPILib.getEquipementHandle(TC_BUS)
        az = OlxAPILib.getEquipementData(busHnd,[BUS_nArea,BUS_nZone],VT_INTEGER)
        self.area = list(set(az[0]))
        self.zone = list(set(az[1]))
        #
        self.initGUI()
    #
    def initGUI(self):
        olrFile = os.path.basename(sys.argv[1])
        self.parent.wm_title("Change bus: "+olrFile)
        try:
            self.parent.wm_iconbitmap(OLXAPI_DLL_PATH+"aspen.ico")
        except:
            pass
        #
        la1 = tk.Label(self.parent, text="Change All")
        la1.pack()
        la1.place(x=20, y=20)
        #
        self.var1 = tk.IntVar()
        #
        self.r1=tk.Radiobutton(self.parent, text="Area",variable = self.var1,value=1, command=self.getValArea)
        self.r1.pack(anchor=tk.W)
        self.r2=tk.Radiobutton(self.parent, text="Zone",variable = self.var1,value=2, command=self.getValZone)
        self.r2.pack(anchor=tk.W)
        self.r1.place(x=100,y=20)
        self.r2.place(x=200,y=20)
        #
        self.az = "Area"
        self.r1.select()
        #
        la2 = tk.Label(self.parent, text="From                         To")
        la2.pack()
        la2.place(x=100, y=70)
        #
        self.updateCombo()
        #
        self.sto = tk.Text(self.parent, height=1, width=6)
        self.sto.pack()
        self.sto.place(x=200, y=90)
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
        self.cb.place(x=100, y=90)
        self.cb.bind("<<ComboboxSelected>>")
        self.cb.current(0)
    #
    def getValArea(self):
        self.az = "Area"
        print("\tselected: "+self.az)
        self.updateCombo()
    #
    def getValZone(self):
        self.az = "Zone"
        print("\tselected: "+self.az)
        self.updateCombo()
    #
    def saveAsOlr(self):
        fileNew = tkFileDialog.asksaveasfilename(defaultextension="OLR",filetypes=[("Python Files", "*.OLR"), ("All Files", "*.*")])
        fileNew = fileNew.replace("/","\\")
        if fileNew:
            if OLRXAPI_FAILURE == OlxAPI.SaveDataFile(fileNew):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            return fileNew
        return ''
    #
    def run_OK(self):
        valFrom = int(self.cb.get())
        try:
            valNew = int(self.sto.get("1.0",tk.END))
        except:
            tkMessageBox.showerror("ERROR","Value To must have INTEGER type")
            return
        #
        if self.az=="Area":
            nParamCode = BUS_nArea
        elif self.az=="Zone":
            nParamCode = BUS_nZone
        #
        kd = changebus(nParamCode,valFrom,valNew)
        if kd ==0:
            tkMessageBox.showinfo("","No bus changed")
        else:
            s2 = str(kd) +" bus have changed "+self.az+ " ("+ str(valFrom) + "=>"+ str(valNew) + ")"
            print(s2)
            fileNew = self.saveAsOlr()
            if fileNew:
                print("File Save As: ",fileNew)
                tkMessageBox.showinfo("",s2+"\n\n"+fileNew+"\n\thad been saved successfully")
        #
        open_olrFile(readonly=0)
# CORE
def changebus(nParamCode,valOld,valNew):
    """
    Change all bus data
        Args :
            nParamCode : BUS_nArea/BUS_nZone
            valOld,valNew : valOld => valNew
        Returns:
            NONE
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






