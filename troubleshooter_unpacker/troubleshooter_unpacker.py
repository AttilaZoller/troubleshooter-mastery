import os
import shutil
import re
import threading
import zipfile
import time

ZIP_SIGNATURE = b"PK\x03\x04\x14\x00\x00"
PNG_SIGNATURE = b"\x89\x50\x4E\x47"
FBX_SIGNATURE = b"\x4B\x61\x79\x64\x61\x72\x61\x20\x46\x42\x58"

lock = threading.Lock()
thread_limiter = threading.Semaphore(50)

def main():
    file_list = []
    for (root, dirs, files) in os.walk("./"):
        if len(files) > 0:
            for filename in files:
                file_list.append({"dir":root,"filename":filename})

    for file_info in file_list:
        thread_limiter.acquire()
        thread = threading.Thread(target=check_filetype, args=(file_info,))
        print(file_info)
        thread.start()
        thread.join()
        thread_limiter.release()


def check_filetype(file_info):
    filename = file_info["dir"] + "/" + file_info["filename"]

    if ".py" in file_info["filename"] or "errorlog" in file_info["filename"]:
        return

    with open(filename,mode="rb") as f:
        byte = f.read(22)

    if ZIP_SIGNATURE in byte: # ZIP 파일이라면, 압축을 풀고, 기존 zip은 삭제한 뒤 한 번 더 체크한다.
        lock.acquire()
        new_filename = filename.replace(filename, filename + ".zip")
        os.rename(filename, new_filename)

        with zipfile.ZipFile(new_filename) as zf:
            zf.extractall(file_info["dir"])

        check_filetype(file_info)
        os.remove(new_filename)
        lock.release()

    elif PNG_SIGNATURE in byte:
        lock.acquire()
        move_file(file_info,"png")
        lock.release()

    elif FBX_SIGNATURE in byte:
        lock.acquire()
        move_file(file_info,"fbx")
        lock.release()

    else:
        try:
            lock.acquire()
            byte_info = str(byte[:4]).replace("\\x","").replace("b","").replace("\'","")
            with open("byte.csv","a") as f:
                f.write("{},{}\n".format(byte_info,file_info["filename"]))
        except:
            with open("errorlog.txt","a") as f:
                f.write(file_info["filename"])
                f.write("\n")
        finally:
            lock.release()

def move_file(file_info,file_extension):
    if not os.path.isdir(file_extension):
        os.mkdir(file_extension)


    filename = file_info["dir"] + "/" + file_info["filename"]

    file_info["filename"] = re.sub(r"(\.{})+$".format(file_extension),"",file_info["filename"])

    if os.path.isfile("./{}/{}.{}".format(file_extension,file_info["filename"],file_extension)):
        count = 1
        while True:
            if os.path.isfile("./{}/{} ({}).{}".format(file_extension,file_info["filename"],str(count),file_extension)):
                count = count + 1
                continue
            else:
                shutil.move(filename,"./{}/{} ({}).{}".format(file_extension,file_info["filename"],str(count),file_extension))
                break
    else:
        shutil.move(filename,"./{}/{}.{}".format(file_extension,file_info["filename"],file_extension))

def make_log(byte,filename):
    with open("filetype_log.csv","a",newline="") as f:
        f.write("{},{}\n".format(byte[:4],filename))

main()