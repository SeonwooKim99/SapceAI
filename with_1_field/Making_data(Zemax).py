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


desktop_path = 'C:/Users/galaxy/Desktop/data_for_AI'

for i in range(10000) :

#Make random errors
    random_values = [round(random.uniform(-5, 5),3) for _ in range(5)]  
    random_thickness = round(random.uniform(-5, 5),3)



    surf11 = TheSystem.LDE.GetSurfaceAt(11)
    surf11.Thickness = random_thickness

    surf10 = TheSystem.LDE.GetSurfaceAt(10)
    for j in range(5) :
        par = j + 12
        surf10.GetCellAt(par).DoubleValue = random_values[j]

    random_values.insert(2,random_thickness)
    

#Make Path(new folder) for saving data
    folder_name = 'data' + str(i+1)
    save_path = os.path.join(desktop_path, folder_name)
    os.makedirs(save_path, exist_ok=True)

#에러값(정답데이터) 저장
    txt_file_path=os.path.join(save_path,"error.txt")
    with open(txt_file_path,"w") as file :
        for item in random_values :
            file.write(f"{item}\n")

#Make wavefront map(TXT) & Save the file
    wfm = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.WavefrontMap)
    wfm_setting = wfm.GetSettings()
    #wfm_setting.Field.SetFieldNumber(2)
    #wfm_setting.Sampling = ZOSAPI.Analysis.SampleSizes.S_128x128
    wfm_setting.RemoveTilt = True
    wfm.ApplyAndWaitForCompletion()
    wfm_result = wfm.GetResults()
    #Path = 'C:/Users/galaxy/Desktop/김선우/wfm.txt'
    wfm_file_path=os.path.join(save_path,"wfm.txt")
    wfm_result.GetTextFile(wfm_file_path)
    wfm.Close()
    

#Zernike Coefficient
    Zernike = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.ZernikeFringeCoefficients)
    Zernike_setting = Zernike.GetSettings()
    #Zernike_setting.Field.SetFieldNumber(2)
    #Zernike_setting.Sampling = ZOSAPI.Analysis.SampleSizes.S_128x128
    Zernike.ApplyAndWaitForCompletion()
    Zernike_result = Zernike.GetResults()


    Zernike_file_path=os.path.join(save_path,"Zernike.txt")
    Zernike_result.GetTextFile(Zernike_file_path)
    Zernike.Close()



    with open (Zernike_file_path, 'r', encoding = 'utf-16') as file :
        lines = file.readlines()

    z5_z9_lines = lines[34:39]

    z5_z9 = [line.split()[2] for line in z5_z9_lines]

    with open(Zernike_file_path, 'w', encoding='utf-16') as output :
        for coeff in z5_z9 :
            output.write(coeff + '\n')
    
        
    

#부여 에러 초기화
    surf11.Thickness = 0  #dz
    for k in range(5) :   #dx,dy,tx,ty,tz
        par = k + 12
        surf10.GetCellAt(par).DoubleValue = 0   
    
    print(i+1, "clear")

print("Completely Save all the Files!")

