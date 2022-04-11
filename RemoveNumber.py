import sys
import os
def removeNumber(path):
    #file_paths = sys.argv[1:]  # the first argument is the script itself
    #print(file_paths[0][-2]=="\")  
    """
    for i in range(len(file_paths[0])-1,0,-1):
        #print(file_paths[0][i])
        
        if file_paths[0][i]=="\\":
            [0:i+1]
            break
            
    """
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
        newname=path+ os.sep + fileList[i]
        if len(fileList[i])>2:
            for j in range(len(fileList[i])):
                if fileList[i][j]=='_':
                    newname=path+ os.sep + fileList[i][j+1:]
                    break
                    
        
        #设置新文件名
        #newname=path + os.sep +s+'_'+fileList[i]
        
        os.rename(oldname,newname)   #用os模块中的rename方法对文件改名
        #print(oldname,'======>',newname)
    