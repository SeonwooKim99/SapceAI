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


desktop_path = 'D:/AI_Alignment/algorithm_data'

for i in range(20000) :

#Make random errors
    random_values_decenter = [round(random.uniform(-0.6, 0.6),3) for _ in range(2)]  
    random_thickness = round(random.uniform(-0.6, 0.6),3)
    random_values_tilt = [round(random.uniform(-0.6, 0.6),3) for _ in range(3)]



    surf13 = TheSystem.LDE.GetSurfaceAt(13)
    surf13.Thickness = random_thickness

# Insert Error
    # Decenter Z
    surf12 = TheSystem.LDE.GetSurfaceAt(12)
    # Dencenter X,Y
    for j in range(2) :
        par = j + 12
        surf12.GetCellAt(par).DoubleValue = random_values_decenter[j]
    # Tilt X,Y,Z
    for k in range(3) :
        par = k + 14   
        surf12.GetCellAt(par).DoubleValue = random_values_tilt[k]
    
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
            
    
    ## Get Wavefront map & RMSE
    RMS = []
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
        
        #RMSE 정보만 저장
        file_first = open(wfm_path,'r', encoding = 'utf-16')
        rmse = file_first.read().splitlines()[9].split()[8]
        RMS.append(float(rmse))

        # wavefront map 정보만 저장
        with open(wfm_path, 'r+', encoding='utf-16') as file:
            lines = file.readlines()
            # 17번째 줄부터 wavefront map 정보
            lines_to_save = lines[16:]
    
            # 파일 수정 후 저장
            file.seek(0)
            file.truncate()
            file.writelines(lines_to_save)

    # RMSE 저장
    RMS_file_path=os.path.join(save_path,"RMSE_with_error.txt")
    with open(RMS_file_path,"w") as file :
        for item in RMS :
            file.write(f"{item}\n")

        
        
    #Zernike coefficient
    for fd in range(5) :
        fdn = fd + 1
        
        zern = TheSystem.Analyses.New_ZernikeFringeCoefficients()
        zern_setting = zern.GetSettings()
        zern_setting = ZOSAPI.Analysis.Settings.Aberrations.IAS_ZernikeFringeCoefficients(zern_setting)

        zern_setting.Field.SetFieldNumber(fdn)
        zern.ApplyAndWaitForCompletion()
        zern_results = zern.GetResults()
        file_name_zer = 'zer_' + str(fdn) + '.txt'
        zer_path = os.path.join(save_path,file_name_zer)
        zern_results.GetTextFile(zer_path) #Result file
        zern.Close()
        
        with open(zer_path, 'r', encoding='utf-16') as zer_file:
            lines = zer_file.read().splitlines()

        zernike = []
        for zer_num in range(34, 39):
            data = lines[zer_num].split()[2]
            zernike.append(data)

        # zernike 리스트의 데이터를 zer_path 파일에 덮어쓰기합니다.
        with open(zer_path, 'w', encoding='utf-16') as zer_file:
            for item in zernike:
                zer_file.write(item + '\n')
            
            
        
        
    
    '''
    TheMFE = TheSystem.MFE

    MaxTerm = 9


    ZERNType = ZOSAPI.Editors.MFE.MeritOperandType.ZERN


    FirstOperand = TheMFE.GetOperandAt(1)


    FirstOperand.ChangeType(ZERNType)



    Wave = 1
    Samp = 2
    #Field = 2
    Type = 0 # Fringe coefficients, 1 = Standard coefficient
    Epsilon = 0.0
    Vertex = 0

    for f in range(5) :
        Field = f + 1
        FirstOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue = MaxTerm     # Term
        FirstOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param2).IntegerValue = Wave        # Wave
        FirstOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param3).IntegerValue = Samp        # Samp
        FirstOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param4).IntegerValue = Field       # Field
        FirstOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param5).IntegerValue = Type        # Type
        FirstOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param6).DoubleValue = Epsilon      # Epsilon
        FirstOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param7).IntegerValue = Vertex      # Vertex

            # Insert new operands and set their values
        for Term in range(1, MaxTerm):
            NewOperand = TheMFE.InsertNewOperandAt(2)
            NewOperand.ChangeType(ZERNType)

            # Adjust its settings with explicit enum conversion
            NewOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue = Term     # Term
            NewOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param2).IntegerValue = Wave     # Wave
            NewOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param3).IntegerValue = Samp     # Samp
            NewOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param4).IntegerValue = Field    # Field
            NewOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param5).IntegerValue = Type     # Type
            NewOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param6).DoubleValue = Epsilon   # Epsilon
            NewOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param7).IntegerValue = Vertex   # Vertex


        TheMFE.CalculateMeritFunction()


        coefficient = []
        for Index in range(MaxTerm, 0, -1):

            CurrentOperand = TheMFE.GetOperandAt(Index)
            Term = CurrentOperand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue
            TermValue = CurrentOperand.Value
            #print(round(TermValue,8))
            coefficient.append(round(TermValue,8))
        z5_9 = coefficient[4:]
        file_name_zer = 'Zernike_' + str(Field) + '.txt'
        zer_path = os.path.join(save_path,file_name_zer)
        with open(zer_path,'w') as file :
            for item in z5_9 :
                file.write(f"{item}\n")
    '''
            
            
        
        
        #부여 에러 초기화
    surf13.Thickness = 0  #dz
    for k in range(5) :   #dx,dy,tx,ty,tz
        par = k + 12
        surf12.GetCellAt(par).DoubleValue = 0   
    
    print(i+1, "clear")


print('Complete get data! Make your AI algorithm! XD')
