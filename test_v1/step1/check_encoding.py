
import chardet

file_path = 'd:/Workspace/Manage_TO/Match_Table1.csv'
try:
    with open(file_path, 'rb') as f:
        rawdata = f.read(1000)
    result = chardet.detect(rawdata)
    print(result)
except Exception as e:
    print(e)
