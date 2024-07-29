import clr, os, winreg
from itertools import islice

# This boilerplate requires the 'pythonnet' module.
# The following instructions are for installing the 'pythonnet' module via pip:
#    1. Ensure you are running a Python version compatible with PythonNET. Check the article "ZOS-API using Python.NET" or
#    "Getting started with Python" in our knowledge base for more details.
#    2. Install 'pythonnet' from pip via a command prompt (type 'cmd' from the start menu or press Windows + R and type 'cmd' then enter)
#
#        python -m pip install pythonnet

# determine the Zemax working directory
aKey = winreg.OpenKey(winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER), r"Software\Zemax", 0, winreg.KEY_READ)
zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
winreg.CloseKey(aKey)

# add the NetHelper DLL for locating the OpticStudio install folder
clr.AddReference(NetHelper)
import ZOSAPI_NetHelper

pathToInstall = ''
# uncomment the following line to use a specific instance of the ZOS-API assemblies
#pathToInstall = r'C:\C:\Program Files\Zemax OpticStudio'

# connect to OpticStudio
success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(pathToInstall);

zemaxDir = ''
if success:
    zemaxDir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory();
    print('Found OpticStudio at:   %s' + zemaxDir);
else:
    raise Exception('Cannot find OpticStudio')

# load the ZOS-API assemblies
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI.dll'))
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI_Interfaces.dll'))
import ZOSAPI

TheConnection = ZOSAPI.ZOSAPI_Connection()
if TheConnection is None:
    raise Exception("Unable to intialize NET connection to ZOSAPI")

TheApplication = TheConnection.ConnectAsExtension(0)
if TheApplication is None:
    raise Exception("Unable to acquire ZOSAPI application")

if TheApplication.IsValidLicenseForAPI == False:
    raise Exception("License is not valid for ZOSAPI use.  Make sure you have enabled 'Programming > Interactive Extension' from the OpticStudio GUI.")

TheSystem = TheApplication.PrimarySystem
if TheSystem is None:
    raise Exception("Unable to acquire Primary system")

def reshape(data, x, y, transpose = False):
    """Converts a System.Double[,] to a 2D list for plotting or post processing
    
    Parameters
    ----------
    data      : System.Double[,] data directly from ZOS-API 
    x         : x width of new 2D list [use var.GetLength(0) for dimension]
    y         : y width of new 2D list [use var.GetLength(1) for dimension]
    transpose : transposes data; needed for some multi-dimensional line series data
    
    Returns
    -------
    res       : 2D list; can be directly used with Matplotlib or converted to
                a numpy array using numpy.asarray(res)
    """
    if type(data) is not list:
        data = list(data)
    var_lst = [y] * x;
    it = iter(data)
    res = [list(islice(it, i)) for i in var_lst]
    if transpose:
        return self.transpose(res);
    return res
    
def transpose(data):
    """Transposes a 2D list (Python3.x or greater).  
    
    Useful for converting mutli-dimensional line series (i.e. FFT PSF)
    
    Parameters
    ----------
    data      : Python native list (if using System.Data[,] object reshape first)    
    
    Returns
    -------
    res       : transposed 2D list
    """
    if type(data) is not list:
        data = list(data)
    return list(map(list, zip(*data)))

print('Connected to OpticStudio')

# The connection should now be ready to use.  For example:
print('Serial #: ', TheApplication.SerialCode)

# Insert Code Here

import numpy as np
import matplotlib.pyplot as plt
import random
import shutil
from enum import Enum
import os


desktop_path = 'D:/AI_Alignment/test'

for i in range(10) :

#Make random errors
    random_values_decenter = [round(random.uniform(-0.5, 0.5),3) for _ in range(2)]  
    random_thickness = round(random.uniform(-0.5, 0.5),3)
    random_values_tilt = [round(random.uniform(-0.3, 0.3),3) for _ in range(3)]



    surf11 = TheSystem.LDE.GetSurfaceAt(11)
    surf11.Thickness = random_thickness

# Insert Error
    # Decenter Z
    surf10 = TheSystem.LDE.GetSurfaceAt(10)
    # Dencenter X,Y
    for j in range(2) :
        par = j + 12
        surf10.GetCellAt(par).DoubleValue = random_values_decenter[j]
    # Tilt X,Y,Z
    for k in range(3) :
        par = k + 14   
        surf10.GetCellAt(par).DoubleValue = random_values_tilt[k]
    
    random_values_decenter.append(random_thickness)
    random_error = random_values_decenter + random_values_tilt
    
    #Make Path(new folder) for saving data
    folder_name = 'data' + str(i+1)
    save_path = os.path.join(desktop_path, folder_name)
    os.makedirs(save_path, exist_ok=True)
    
    #Save error data
    txt_file_path=os.path.join(save_path,"error.txt")
    with open(txt_file_path,"w") as file :
        for item in random_error :
            file.write(f"{item}\n")
            
    # Get Wavefrontmap, Zernike coefficient & Save file
    for field in range(5) :
        filed_number = field + 1
        
        # WFM
        Mywavefront = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.WavefrontMap)
        Mywavefront.WaitForCompletion()
        Mywavefront_settings = Mywavefront.GetSettings()

        Mywavefront_settings = ZOSAPI.Analysis.Settings.IAS_WavefrontMap(Mywavefront_settings)

    
        Mywavefront_settings.Field.SetFieldNumber(filed_number) #Field 설정
        Mywavefront_settings.RemoveTilt = True

        #RemoveTilt 설정
        Mywavefront.ApplyAndWaitForCompletion()
        Mywavefront_result = Mywavefront.GetResults()
        file_name_wfm = 'wfm_' + str(filed_number) + '.txt'
        wfm_path = os.path.join(save_path,file_name_wfm)
        Mywavefront_result.GetTextFile(wfm_path)#Result file 추출
        Mywavefront.Close()
        
        
        #Zernike coefficient
        """
        Zernike = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.ZernikeFringeCoefficients)
        Zernike_setting = Zernike.GetSettings()
        Zernike_setting.Field.SetFieldNumber(filed_number)
        Zernike.ApplyAndWaitForCompletion()
        Zernike_result = Zernike.GetResults()

        file_name_Zernike = 'Zernike_' + str(filed_number) + '.txt'
        Zernike_file_path=os.path.join(save_path,file_name_Zernike)
        Zernike_result.GetTextFile(Zernike_file_path)
        Zernike.Close()
        """
        #부여 에러 초기화
    surf11.Thickness = 0  #dz
    for k in range(5) :   #dx,dy,tx,ty,tz
        par = k + 12
        surf10.GetCellAt(par).DoubleValue = 0   
    
    print(i+1, "clear")

    
