import os

def input_fiozp():
    fiozp_load = []
    dir_path = os.path.abspath('xls_files/')
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            if file.endswith('.xls'):
                fiozp_load.append(file)
    return fiozp_load

if __name__ == '__main__':
    print(input_fiozp())
    print(os.path.abspath('xls_files/'))