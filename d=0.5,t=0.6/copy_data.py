import os
import shutil

# 원본 폴더 경로와 대상 폴더 경로 설정
source_base_path = 'D:/AI_Alignment/algorithm_data'
target_base_path = 'D:/AI_Alignment/data_second'

# 복사할 파일 리스트
files_to_copy = [
    'wfm_1.txt', 'wfm_2.txt', 'wfm_3.txt', 'wfm_4.txt', 'wfm_5.txt', 
    'error.txt', 'RMSE_with_error.txt'
]

# 원본 폴더의 번호 범위 설정
folder_count = 20000

# 각 폴더에 대해 작업 수행
for i in range(1, folder_count + 1):
    # 원본 폴더 경로 설정
    source_folder = os.path.join(source_base_path, f'data{i}')
    
    # 대상 폴더 경로 설정
    target_folder = os.path.join(target_base_path, f'data{i}')
    
    # 대상 폴더가 존재하지 않으면 생성
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # 복사할 파일들에 대해 작업 수행
    for file_name in files_to_copy:
        source_file = os.path.join(source_folder, file_name)
        target_file = os.path.join(target_folder, file_name)
        
        # 파일이 존재하면 복사
        if os.path.exists(source_file):
            shutil.copy(source_file, target_file)

print("모든 파일 복사가 완료되었습니다.")

