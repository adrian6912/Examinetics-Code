import dicom as d
import os

def main():
    names = dict()
    g = 'F:'
    AccNum = input('Enter Accession Number')
    f = []
    
    os.chdir(g)

    for (dirpath, dirnames, filenames) in os.walk(g):
        f.extend(filenames)
        print(f)
        break
    
    for file in f:
        try:
            if '.dcm' in file:
                plan = d.read_file(file)
                print(plan.AccessionNumber)
                if plan.AccessionNumber == AccNum:
                    pass
                else:
                    plan.AccessionNumber = AccNum
                    plan.save_as(file)
            elif '.DCM' in file:
                plan = d.read_file(file)
                print(plan.AccessionNumber)
                if plan.AccessionNumber == AccNum:
                    pass
                else:
                    plan.AccessionNumber = AccNum
                    plan.save_as(file)
            else:
                pass
        except d.errors.InvalidDicomError:
            print("Invalid file. Name: " + file)


if __name__ == '__main__':
    main()
