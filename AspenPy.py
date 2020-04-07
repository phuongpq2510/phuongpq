"""
Python path link to python,OlxAPI,OlxAPILib,OlxAPIConst
    System parameters:
        sys.argv[1] = (str) OLR file
        sys.argv[2] = (Int 1/0) loaded DLL

"""
__author__ = "ASPEN"
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "15.1.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

#
PYTHON2_PATH       = "C:\\Python27\\"                                          # Python 27
OLXAPI_DLL_PATH    = "C:\\Program Files (x86)\\ASPEN\\1LPFv15\\"               # whete located olxapi.dll
OLXAPI_PYTHON_PATH = "C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python" # where located OlxAPI,OlxAPILib,OlxAPIConst
OLR_DEFAULT_PATH   = "C:\\Program Files (x86)\\ASPEN\\1LPFv15\\SAMPLE30.OLR"   # default example
##OLR_DEFAULT_PATH   = "D:\\NLDC\\perso\\apo\\Dev\\SAMPLE30.OLR"

#
import sys,os
sys.path.append(OLXAPI_PYTHON_PATH)
os.environ['PATH'] = OLXAPI_PYTHON_PATH + ";" + os.environ['PATH']
#
import OlxAPI
import OlxAPILib
from OlxAPIConst import *
from ctypes import *
#
def load_olxapi_dll():
    """
    Load olxapi.dll with test in
        sys.argv[2] = (Int 1/0) loaded DLL
    """
    try:
        loadDLL = int(sys.argv[2])
    except:
        setPara_sys_argv(idx=2,val=0)
    loadDLL = int(sys.argv[2])
    #
    if loadDLL==0:
        OlxAPI.InitOlxAPI(OLXAPI_DLL_PATH)
        setPara_sys_argv(idx=2,val=1)
        print ("\tOlxAPI : Version:"+str(OlxAPI.Version())+" Build: "+str(OlxAPI.BuildNumber()))
#
def open_olrFile(readonly):
    """
    Open OLR file in
        sys.argv[1] = (str) OLR file

    Args:
        readonly (int): open in read-only mode. 1-true; 0-false
    """
    try:
        olrFile = sys.argv[1]
    except:
        setPara_sys_argv(idx=1,val=OLR_DEFAULT_PATH)
    olrFile = sys.argv[1]
    #
    if OLRXAPI_OK == OlxAPI.LoadDataFile(olrFile,readonly):
         print( "\tFile opened successfully: " + str(olrFile))
    else:
         raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
#
def setPara_sys_argv(idx,val):
    """
    set
        sys.argv[idx] = val
    """
    for i in range(len(sys.argv),idx+1):
        sys.argv.append(None)
    sys.argv[idx] = val



