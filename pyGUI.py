"""
Demo usage of ASPEN OlrxAPI in Python.

Purpose: GUI for show and run python in a directory
"""
from __future__ import print_function

__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.0.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

#
from AspenPy import*
#
import Tkinter as tk
import tkFileDialog,tkMessageBox
import sys,os
import runpy,subprocess

class APP(tk.Frame):
    def __init__(self, parent,olrFile,dirPy):
        self.parent = parent
        self.dirPy = dirPy
        self.olrFile = olrFile
        #
        self.initGUI()
    #
    def initGUI(self):
        self.parent.wm_title("PYTHON TOOL")
        try:
            self.parent.wm_iconbitmap(OLXAPI_DLL_PATH+"aspen.ico")
        except:
            pass

        self.parent.config(width = 800, height = 400,background="lavender")
        # MENU -----------------------------------------------------------------
        menubar = tk.Menu(self.parent)
        filemenu = tk.Menu(menubar, tearoff=0)
        #
        filemenu.add_command(label="Select OLR file", command=self.open_olrfile)
        filemenu.add_command(label="Change python directory", command=self.open_dir)
        filemenu.add_separator()
        filemenu.add_command(label="Save selected python"   , command=self.save_py)
        filemenu.add_command(label="Save As selected python", command=self.saveas_py)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.parent.destroy)
        menubar.add_cascade(label="File", menu=filemenu)
        #
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About...", command=self.help_about)
        menubar.add_cascade(label="Help", menu=helpmenu)
        self.parent.config(menu=menubar)

        # FILE FRAME
        frame_file = tk.Frame(self.parent, relief=tk.RAISED, bd=2,bg= 'lavender')
        #
        button_olr = tk.Button(frame_file, text="Olr file", command=self.open_olrfile)
        button_olr.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        #
        self.text_olr = tk.Text(frame_file, height=1, width=25,bg='moccasin')
        self.text_olr.grid(row=1, column=0, sticky="ew", padx=3, pady=3)

        #
        button_dir = tk.Button(frame_file, text="Python directory", command=self.open_dir)
        button_dir.grid(row=2, column=0,sticky="ew",  padx=3, pady=3)
        #
        self.text_dir = tk.Text(frame_file, height=1, width=25,bg='bisque')
        self.text_dir.grid(row=3, column=0, sticky="ew", padx=3, pady=3)

        #
        self.lbox = tk.Listbox(frame_file,height=20)
        self.lbox.grid(row=4, column=0,sticky="ew",  padx=5, pady=5)
        self.lbox.bind("<<ListboxSelect>>", self.showcontent)

        # RIGHT FRAME
        frame_right = tk.Frame(self.parent, relief=tk.RAISED, bd=2,bg= 'lavender')
        #
        frame_pyb = tk.Frame(frame_right, relief=tk.RAISED, bd=2,bg= 'lavender')
        frame_pyb.grid(row=0, column=0, sticky="ew", padx=3, pady=3)
        #button run
        button_run = tk.Button(frame_pyb , text = ' Run selected python ',command=self.runPy)
        button_run.grid(row=0, column=0, sticky="ew", padx=3, pady=3)

        #button edit
        button_edit = tk.Button(frame_pyb , text = ' Edit  selected python ',command=self.editPy)
        button_edit.grid(row=0, column=2, sticky="ew", padx=3, pady=3)

         #button edit
        button_edit = tk.Button(frame_pyb , text = ' Save  selected python ',command=self.save_py)
        button_edit.grid(row=0, column=3, sticky="ew", padx=3, pady=3)

        # Frame content Py
        frame_py = tk.Frame(frame_right, relief=tk.RAISED, bd=2,bg= 'lavender')
        frame_py.grid(row=1, column=0, sticky="ew", padx=3, pady=3)

        # py_edit
        self.py_edit = tk.Text(frame_py,wrap = tk.NONE,height = 25, width = 68)
        self.py_edit.grid(row=1, column=0,sticky="ew",  padx=3, pady=3)#
        # yScroll py_edit
        yscroll = tk.Scrollbar(frame_py, orient=tk.VERTICAL, command=self.py_edit.yview)
        self.py_edit['yscroll'] = yscroll.set
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.py_edit.pack(fill=tk.BOTH, expand=tk.Y)
        # xScroll py_edit
        xscroll = tk.Scrollbar(frame_py, orient=tk.HORIZONTAL, command=self.py_edit.xview)
        self.py_edit['xscroll'] = xscroll.set
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.py_edit.pack(fill=tk.BOTH, expand=tk.Y)
        #
        self.setOlrFile(self.olrFile)
        self.setDirPy(self.dirPy)
        self.showDirPy()
        self.showPy()
                #
        frame_file.grid(row=0, column=0, sticky="ns")
        frame_right.grid(row=0, column=1, sticky="ns")
     #
    def setOlrFile(self,olrFile):
        if os.path.isfile(olrFile):
            self.olrFile = olrFile
            #
            self.text_olr.delete(1.0, tk.END)
            self.text_olr.insert(tk.END, self.getNameFileReduc(olrFile))
        else:
            self.olrFile = ""
    #
    def getNameFileReduc(self,f0):
        nmax = 23
        if len(f0)<nmax:
            return f0
        #
        f1= str(f0).replace("\\","/")
        f1= f1.replace("\"","/")
        #
        idx= [pos for pos, char in enumerate(f1) if char == "/"]
        for i in idx:
            if len(f1)-i<nmax:
                return ".."+f1[i:]
        #
        return f1
    #
    def setDirPy(self,dirPy):
        if os.path.isdir(dirPy):
            self.dirPy= dirPy
            #
            self.text_dir.delete(1.0, tk.END)
            s = self.getNameFileReduc(dirPy)
            self.text_dir.insert(tk.END, s)
        else:
            self.dirPy = ""

    #
    def open_olrfile(self):
        """Open a file for editing."""
        olrFile = tkFileDialog.askopenfilename(filetypes=[("Olr Files", "*.olr"), ("All Files", "*.*")])
        if not olrFile:
            olrFile = ""
        #
        self.setOlrFile(olrFile)
        #
        print("Select File: ",olrFile)

    #
    def open_dir(self):
        """Open a directory."""
        self.dirPy = tkFileDialog.askdirectory()
        self.setDirPy(self.dirPy)
        self.showDirPy()
    #
    def saveas_py(self):
        """Save the current file as a new file."""
        filepath = tkFileDialog.asksaveasfilename(defaultextension="py",filetypes=[("Python Files", "*.py"), ("All Files", "*.*")],)
        if not filepath:
            return
        with open(filepath, "w") as output_file:
            text = self.py_edit.get(1.0, tk.END)
            output_file.write(text)
        #
        s1 = filepath + "\n\thad been saved successfully"
        print(s1)
        tkMessageBox.showinfo("",s1)

    def save_py(self):
        """Save the current file """
        with open(self.currentPy, "w") as output_file:
            text = self.py_edit.get(1.0, tk.END)
            output_file.write(text)
        print("Save: ",self.currentPy)
    #
    def runPy(self):
        """
        Run select python
        BUG here
        """
        if not os.path.isfile(self.currentPy):
            tkMessageBox.showwarning("Warning","Null selected python")
            return
        #
        if len(sys.argv)==1:
            sys.argv.append(self.olrFile)
        else:
            sys.argv[1]= self.olrFile
        #
        runpy.run_path(self.currentPy, run_name='__main__')
##        runpy.run_path(self.currentPy)
##        execfile(self.currentPy)
##        subprocess.Popen(self.currentPy, shell=True,stdout=subprocess.PIPE, stderr=subprocess.PIPE )
##        for line in process.stdout:
##            result.append(line)
##        errcode = process.returncode
##        for line in result:
##            print(line)
        #
        print("\tFinish run file:",self.currentPy)
    #
    def showPy(self):
        if os.path.isfile(self.currentPy):
            with open(self.currentPy) as file:
                f1 = file.read()
                self.py_edit.delete('1.0', tk.END)
            self.py_edit.insert(tk.END, f1)
    #
    def showDirPy(self):
        #
        flist = []
        print("dirPy",self.dirPy)
        if os.path.isdir(self.dirPy):
            flist1 = os.listdir(self.dirPy)
            for item in flist1:
                if str(item).lower().endswith(".py"):
                    flist.append(item)
        #
        self.lbox.delete(0, tk.END)
        #
        for item in flist:
            self.lbox.insert(tk.END, item)
        #
        self.currentPy = ""


    def showcontent(self,event):#event
        x = self.lbox.curselection()[0]
        self.currentPy = self.dirPy + "\\" +self.lbox.get(x)
        self.showPy()
    #
    def help_about(self):
        print("Waiting to add")
    #
    def editPy(self):
        if not os.path.isfile(self.currentPy):
            tkMessageBox.showwarning("Warning","Null selected python")
            return
        app = PYTHON2_PATH + "Lib\\idlelib\\idle.bat"
        args = [app,self.currentPy]
        subprocess.call(args)

#
if __name__ == "__main__":
    """
    Run with input arguments
        sys.argv[1] = (str) OLR file
        sys.argv[2] = (Int 1/0) loaded DLL
        sys.argv[3] = (str) dir python
    """
    # test and set default input (if need)
    try:
        olrFile = sys.argv[1]
    except:
        setPara_sys_argv(idx=1,val=OLR_DEFAULT_PATH)
    olrFile = sys.argv[1]
    #
    try:
        dirPy = sys.argv[3]
    except:
        dirPy = "D:\\NLDC\\perso\\apo\\Dev\\pyOlxAPIDev_Phuong" #  OLXAPI_PYTHON_PATH "D:\NLDC\perso\apo\Dev\pyOlxAPIDev_Phuong"
        setPara_sys_argv(idx=3,val=dirPy)
    dirPy = sys.argv[3]

    # RUN GUI
    root = tk.Tk()
    root.geometry("800x500+100+50")
    d = APP(root, olrFile=sys.argv[1] , dirPy = sys.argv[3])
    root.mainloop()





