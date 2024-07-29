import numpy as np
import pandas as pd
import random

file_names = ['Align_LSB_1_x+0y+1_LSB_AS_Test1.txt', 'Align_LSB_2_x-1y+0LSB_AS_Test1.txt', 'Align_LSB_3_x+0y+0LSB_AS_Test1.txt', 'Align_LSB_4_x+1y+0LSB_AS_Test1.txt', 'Align_LSB_5_x+0y-1LSB_AS_Test1.txt'] #파일이름


matrices = []

path='C:/Users/galaxy/Desktop/Alignment_Simulation_for_K-DRIFT_Team/Alignment Code/'

for file_name in file_names:
    with open(path+file_name, 'r',encoding='utf-16') as file: # 파일에서 세 번째 줄부터 끝까지의 내용을 읽어서 행렬로 변환
        lines = file.readlines()[2:]  # 세 번째 줄부터 끝까지 읽기
        matrix = np.loadtxt(lines).reshape(-1, 1)  # 행렬로 변환
        matrices.append(matrix)


SM = np.concatenate(matrices)


print("complete merging matrix.")


file_path_IM='C:/Users/galaxy/Desktop/Alignment_Simulation_for_K-DRIFT_Team/Alignment Code/Align_LSB_Cross_Ideal.txt'
file_path_AM='C:/Users/galaxy/Desktop/Alignment_Simulation_for_K-DRIFT_Team/Alignment Code/M2SenCrossZ5Z9.txt'

im=[]
with open(file_path_IM, 'r',encoding='utf-8') as file:
        matrix = np.loadtxt(file).reshape(25, 1)
        im.append(matrix)
IM=np.array(im)

am=[]
with open(file_path_AM, 'r',encoding='utf-8') as file:
        matrix = np.loadtxt(file).reshape(25, 6)
        am.append(matrix)
AM=np.array(am)

DM=SM-IM
DM=DM.reshape(25,1)  # Delta Zernike Coefficient
AM=AM.reshape(25,6)  # Sensitivity Table

def sol(DM, AM) :
  a_t = AM.T
  first= np.linalg.pinv(a_t@AM) 
  x=first@a_t@DM
  return x


Error = sol(DM,AM)

solution = []
for i in range(6) :
     a = round(Error[i][0],4)
     solution.append(a)

print("Complete find Error with Sensitivity Table!")
