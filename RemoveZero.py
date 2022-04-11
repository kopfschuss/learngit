import sys
import os

def removeZero(path):
    #file_paths = sys.argv[1:]  # the first argument is the script itself
    #path=file_paths[0]
    print(os.listdir(path))
    fileList=os.listdir(path)
    fileList.sort()
    for i in range(len(fileList)):
        s=i+1
        s=str(s)
        #s=s.zfill(5)
        
        #设置旧文件名（就是路径+文件名）
        oldname=path+ os.sep + fileList[i]   # os.sep添加系统分隔符
        
        #设置新文件名
        newname=path + os.sep +s+'_'+fileList[i]
        if len(fileList[i])>6 and fileList[i][5]=='_':
            s=str(int(fileList[i][:5]))
            newname=path + os.sep +s+'_'+fileList[i][6:]
        
        os.rename(oldname,newname)   #用os模块中的rename方法对文件改名
        #print(oldname,'======>',newname)
    
