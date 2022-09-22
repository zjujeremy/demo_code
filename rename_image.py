import os
import sys

def rename_image(file_name):
    with open(file_name, "r") as fin:
        c = fin.readlines()
        for line in c:
            index, origin_file, target_file = line.strip().split()
            if os.path.exists(origin_file):
                os.rename(origin_file, target_file)
            else:
                print("{} not exists".format(origin_file))
    print("Done")

if __name__ == "__main__":
    rename_image(sys.argv[1])
