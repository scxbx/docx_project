import os






def get_files():
    sample_f = os.path.join(os.getcwd(), 'summary')
    for filepath,dirnames,filenames in os.walk(sample_f):
        for filename in filenames:
            print(os.path.join(filepath.replace('summary', 'result', 1), filename))


get_files()