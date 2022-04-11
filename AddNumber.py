import os
def addNumber(path,offset):
    offset=int(offset)-1
    print(path,offset)
    fileList=os.listdir(path)
    print(fileList)
    #fileList.sort()
    for i in range(len(fileList)):

        s=i+1
        s=str(s+offset)
        s=s.zfill(5)
        
        #设置旧文件名（就是路径+文件名）
        oldname=path+ os.sep + fileList[i]   # os.sep添加系统分隔符

        
        #设置新文件名
        newname=path + os.sep +s+'_'+fileList[i]
        if len(fileList[i])>2:
            for j in range(len(fileList[i])):
                if fileList[i][j]=='_':
                    if fileList[i][:j].isdigit():
                        newname=path + os.sep +s+'_'+fileList[i][j+1:]
                    #newname=path+ os.sep + fileList[i][j+1:]
                    break
        os.rename(oldname,newname)   #用os模块中的rename方法对文件改名
        #print(oldname,'======>',newname)
        
    #print(path)
    #print(file_paths,file_paths[0])
    """
    for p in file_paths:
        print(p)
    """
def AddNumber():
    import sys
    file_paths = sys.argv[1:]
    path=file_paths[0]
    offset=0
    #path="C:\Users\nwcdi\Downloads\archives\111"
    for i in range(len(path)-1,0,-1):
        if path[i]==os.sep:
            offset=int(path[i+1:])
            break
    offset=offset-1    
    addNumber(path,offset)
