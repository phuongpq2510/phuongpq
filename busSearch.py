"""
Demo usage of ASPEN OlrxAPI in Python.

Purpose: test bus search function

"""
from __future__ import print_function
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__version__   = "0.1.1"
__email__     = "support@aspeninc.com"
__status__    = "In development"

# IMPORT -----------------------------------------------------------------------
import sys,os,time
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
sys.path.insert(0, PATH_LIB)
#
import OlxAPI
import OlxAPILib
from OlxAPIConst import *
# Tkinter python 3|2
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
# IMPORT -----------------------------------------------------------------------

# GUI
class APP(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        #

def run():
    OlxAPILib.load_olxapi_dll()
    ARGVS.fi = os.path.join (PATH_FILE, "SAMPLE300.OLR")
    OlxAPILib.open_olrFile(readonly=1)
    bsear = OlxAPILib.BusSearch(gui=0)
    #
    bhnd = bsear.searchBy_NameKv("Alc",33.5)
##    bhnd = bsear.searchBy_NameKv("v",33.5)
##    bhnd = bsear.searchBy_NameKv("alaskaa",33)
##    bhnd = bsear.searchBy_NameKv("Ph4",503)
    #
##    bhnd = bsear.searchBy_BusNumber(260)



if __name__ == '__main__':
    t0 = time.clock()
    run()
    print("runtime=",round(time.clock()-t0,1))



