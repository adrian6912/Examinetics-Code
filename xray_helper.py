import dicom as d
import os
from openpyxl import load_workbook
import datetime as dt
import shutil

def files_in_folder(temp):
    file_list = []
    
    for (dirpath, dirnames, filenames) in os.walk(temp):
        file_list.extend(filenames)
        break
    
    return file_list

def main():
    names = dict()
    g = 'F:\\1018-31-2017'
    files = []
    dates = dict()
    
    os.chdir(g)

    files = files_in_folder(g)
    excel_files = []
    
    for file in files:
        print("2")
        if '.dcm' in file:
            try:
                plan = d.read_file(file)
                try:
                    dates[plan.StudyDate]
                except KeyError:
                    dates[plan.StudyDate] = [file]
                else:
                    dates[plan.StudyDate].append(file)
            except d.errors.InvalidDicomError:
                print(file)
                pass
        elif '.ods' in file:
            try:
                day = dt.datetime.strptime(str(file.strip('.ods')), '%m%d%Y').strftime('%Y%m%d')
                os.rename(file, day + ".ods")
                excel_files.append(day + '.ods')
            except ValueError:          #If a valueError occurs here, then the ods file names are already in the desired format %Y%m%d
                excel_files.append(file)
    print(excel_files)
    for day in dates.keys():
        try:
            os.makedirs(day)
        except FileExistsError:
            pass
        for file in dates[day]:
            shutil.move(os.getcwd() + "\\" + file, os.getcwd() + "\\" + day + "\\" + file)
            print("DCM " + str(file) + " moved.")
    print("DCMs moved.")
    for file in excel_files:
        shutil.move(os.getcwd() + "\\" + file, os.getcwd() + "\\" + file.strip('.ods') + "\\" + file)
        print("Excel " + str(file) + " moved.")
        


if __name__ == '__main__':
    main()
