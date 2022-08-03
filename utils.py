import os, errno, stat

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory. ' +  directory)

def handleRemoveReadonly(func, path, exc):
    excvalue = exc[1]
    if func in (os.rmdir, os.remove) and excvalue.errno == errno.EACCES:
        os.chmod(path, stat.S_IRWXU| stat.S_IRWXG| stat.S_IRWXO) # 0777
        func(path)
    else:
        raise

def uniq_rename(dest_name, new_name):
    filename, ext = os.path.splitext(new_name)

    uniq = 1
    while os.path.exists(new_name):  #동일한 파일명이 존재할 때
        new_name = f"{filename} ({uniq}){ext}"
        uniq += 1
    
    os.rename(dest_name, new_name)
