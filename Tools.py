import pdftotext
from PIL import Image
from pdf2image import convert_from_path
import pytesseract
import json
import openpyxl
import os

def openExcel(file):
    if os.path.splitext(file)[1]==".xlsm":
        return openpyxl.load_workbook(file,keep_vba=True)
    return openpyxl.load_workbook(file)

def pdfToText(file):
    #return pdf
    return pdftotext.PDF(open(file,'rb'))
def fileToDict(filepath):
    file = open(filepath,'r',encoding='utf-8') 
    js = file.read()
    file.close() 
    return json.loads(js)    
def getDigit(s):
    ret=""
    for i in range(len(s)):
        if s[i].isdigit():
            j=i 
            while j<len(s) and s[j].isdigit():
                j+=1
            # s[j] is not digit, so 
            return int(s[i:j])
    return -1
def imagePdf2text(file,number):
    doc = convert_from_path(file)
    for page_number, page_data in enumerate(doc):
        if page_number==number:
            txt = pytesseract.image_to_string(page_data).encode("utf-8")
            txt=str(txt)
            txt=txt.split("\\n")
            txt=" ".join(txt)
            return txt 
Indication={}#ExtractKeyWords.Indication
CountryDict={}#ExtractKeyWords.CountryDict
def f0(nums):
    
    ans=[]
    #提取方案编号Protocol
    flag=1
    for i in range(len(nums)):
        if nums[i]=='Protocol:':
            ans.append(nums[i+1])
            flag=0
            break
    if flag:
        ans.append("UNK")
    if ans[-1]=="20975":
        ans[-1]="CA209-7JH"
    if ans[-1]=="INCB":
        ans[-1]="CA209-9UJ"        
    if ans[-1]=="NU18C02":
        ans[-1]="CA017-075" 
    if ans[-1]=="PCI-32765LYM1002":
        ans[-1]="CA209-330"             
        
    #查找适应症   
    #d={"CA209-025":"肾透明细胞癌","CA209-026":"IV期或复发性Pd-L1阳性非小细胞肺癌","CA209-032":"晚期或转移性实体瘤","CA209-037":"抗Ctla-4治疗后进展的晚期(不可切除或转移性)黑色素瘤","CA209-039":"复发/难治性血液恶性肿瘤","CA209-040":"晚期肝细胞癌","CA209-066":"未经治疗不可切除或转移性黑色素瘤","CA209-067":"未经治疗的不可切除或转移性黑色素瘤","CA209-077":"晚期或复发性实体瘤","CA209-142":"结肠癌","CA209-153":"晚期或转移性非小细胞肺癌","CA209-171":"非小细胞肺癌","CA209-205":"经典霍奇金淋巴瘤","CA209-227":"化疗-初期IV期或复发性非小细胞肺癌","CA209-238":"高复发风险的3b/c或4期黑色素瘤","CA209-274":"高危侵袭性尿路上皮癌","CA209-275":"转移性或不可切除的尿路上皮癌","CA209-330":"血液肿瘤","CA209-358":"实体瘤","CA209-384":"非小细胞肺癌","CA209-401":"III期(不可切除)或IV期黑色素瘤","CA209-436":"CD30表达的复发难治性非霍奇金淋巴瘤","CA209-451":"广泛性疾病小细胞肺癌","CA209-459":"晚期肝细胞癌","CA209-511":"以前未经治疗不可切除或转移性黑色素瘤","CA209-548":"恶性胶质瘤","CA209-568":"IV期非小细胞肺癌","CA209-577":"切除的食管癌或胃食管交界癌","CA209-592":"生物标记物探索性研究","CA209-596":"复发性GBM","ABTC-1501":"复发性GBM","CA209-602":"复发和难治性多发性骨髓瘤","CA209-627":"晚期或转移性恶性肿瘤","CA209-647":"复发/难治性原发性中枢神经系统淋巴瘤(PCNSL)或复发/难治性原发性睾丸淋巴瘤(PTL)","CA209-648":"食管鳞状细胞癌","CA209-649":"未治疗的晚期或转移性胃或胃食管交界处癌症","CA209-650":"转移性去势抵抗性前列腺癌","CA209-651":"头颈部复发或转移性鳞状细胞癌","CA209-672":"实体瘤","CA209-722":"非小细胞肺癌","CA209-73L":"未经治疗的局限性晚期非小细胞肺癌","CA209-743":"不可切除的胸膜间皮瘤","CA209-74W":"肝细胞癌","CA209-76K":"黑色素瘤","CA209-76U":"恶性黑色素瘤","CA209-7A8":"乳腺癌","CA209-77T":"II-IIIB期非小细胞肺癌","CA209-77W":"转移性去势抵抗性前列腺癌","PICI0033":"转移性去势抵抗性前列腺癌","CA209-7JH":"结直肠癌","CA209-075":"结直肠癌","CA209-7FL":"高危，雌激素受体阳性(ER+)，人表皮生长因子受体2阴性(HER2-)原发性乳腺癌","CA209-7DX":"转移性去势抵抗性前列腺癌","CA209-800":"肾细胞癌","CA209-816":"早期非小细胞肺癌","CA209-817":"非小细胞肺癌","CA209-844":"胃癌术后辅助化疗","ONO-4538-38":"胃癌术后辅助化疗","CA209-848":"晚期或转移性肿瘤突变负荷高的肿瘤","CA209-870":"非小细胞肺癌","CA209-8FC":"黑色素瘤","CA209-8HW":"转移性结直肠癌","CA209-8KX":"多肿瘤药代动力学","CA209-8Y8":"未治疗的肾细胞癌","CA209-8J7":"卵巢癌","CO-338-087":"卵巢癌","CA209-8CH":"晚期乳腺和尿路上皮癌症","DS8201-A-U105":"晚期乳腺和尿路上皮癌症","CA209-901":"未经治疗的不能切除或转移性尿路上皮癌","CA209-914":"局限性肾细胞癌","CA209-915":"黑色素瘤","CA209-9X8":"转移性结直肠癌","CA209-9DW":"晚期肝癌","CA209-9DX":"高的风险肝切除术或消融治疗后复发","CA209-9LA":"非小细胞肺癌","CA209-9ER":"未经治疗的晚期或转移性肾细胞癌","CA209-9KD":"转移性去势抵抗性前列腺癌","CA209-908":"原发性中枢神经系统恶性肿瘤","CA209-9UT":"先前未治疗的晚期或转移性肾细胞癌高风险，非肌肉浸润性膀胱癌","CA209-9UJ":"晚期或转移性恶性肿瘤","INCB24360-208":"晚期或转移性恶性肿瘤","CA209-9F6":"局部晚期或转移性实体肿瘤恶性肿瘤","CA045-013":"局部晚期或转移性实体肿瘤恶性肿瘤","16-214-02":"局部晚期或转移性实体肿瘤恶性肿瘤","CA009-008":"复发胶质母细胞瘤","J17154":"复发胶质母细胞瘤","CA013-004":"晚期实体瘤","CA017-003":"晚期恶性肿瘤","CA017-075":"恶性胶质瘤","CA017-078":"膀胱癌","CA018-003":"晚期胃癌","CA018-005":"晚期肾细胞癌","CA022-001":"晚期实体瘤","CA020-002":"晚期实体瘤","CA025-006":"晚期胰腺癌","CA025-018":"局限性三阴性乳腺癌","CA027-002":"晚期癌症","CA030-001":"晚期实体瘤","CA039-001":"晚期实体瘤","CA045-001":"未经治疗不可切除或转移性黑色素瘤","CA045-002":"未治疗的晚期肾细胞癌","17-214-09":"未治疗的晚期肾细胞癌","CA045-009":"肌层浸润性膀胱癌","CA045-011":"既往未治疗的晚期或转移性肾细胞癌","CA045-012":"晚期或转移性尿路上皮癌症患者","CA045-020":"复发或难治性恶性肿瘤","CA045-022":"黑素瘤","18-214-10":"晚期或转移性尿路上皮癌症患者","CA046-006":"晚期实体瘤","CA223-001":"晚期难治性实体肿瘤","CA224-020":"晚期实体瘤","CA224-047":"未治疗的转移或不可切除的黑色素瘤","CA224-048":"晚期恶性肿瘤","CA224-060":"胃或胃食管交界腺癌","CA224-065":"急性髓系白血病","CA224-073":"晚期肝癌","CA224-083":"转移性恶性黑色素瘤","CA224-104":"复发性非小细胞肺癌","ONO-4538-24E":"食管癌","ONO-4538-37":"不可切除的胃癌","ONO-4538-52":"非鳞状非小细胞肺癌","ONO-4538-53":"晚期或转移性实体瘤","ONO-4538-X53":"上皮癌","ONO-4538-64":"肝细胞癌","ONO-4538-67":"可切除的恶性肿瘤","ONO-4538-74":"子宫颈鳞状细胞癌","ONO-4538-91":"胆道癌","ONO-4578-01":"晚期或转移性实体瘤","SGN35-015":"霍奇金淋巴瘤","SGN35-027":"经典霍奇金淋巴瘤","CV202-103":"晚期实体瘤","17-214-09":"未治疗的晚期肾细胞癌","17-262-01":"局部晚期或转移性实体肿瘤恶性肿瘤","CO-3810-101":"晚期转移性实体瘤","DS8201-A-U105":"晚期乳腺癌和尿路上皮癌","J1714":"II/iIII期食管/胃食管交界癌","XL184-313":"未治疗的中度或低风险的晚期或转移性肾细胞癌","20-214-29":"黑色素瘤","16-214-02":"局部晚期或转移性实体肿瘤恶性肿瘤", "CA224-050": "晚期胃癌", "CA209-67T":"晚期或转移性透明细胞肾细胞癌", "CA043-001": "晚期恶性肿瘤","ONO-4578-05":"晚期或复发性非小细胞肺癌","CA209-7G8":"高危非肌肉浸润性膀胱癌","CA209-920":"未经治疗、晚期或转移性肾细胞癌","CA209-714":"头颈部复发性或转移性鳞状细胞癌","ONO-4538-32":"胆管癌"}
    
    # dict usage: dict[key]=values
    key=ans[0]
    if key in Indication:
        indiction = Indication[ans[0]]
    else:
        indiction = "1"
    ans.append(indiction)        
    return ans

def f1(nums):
    ans=[]

    #提取SUSAR No. 编号
    flag=1
    for i in range(len(nums)):
        if nums[i][:4]=="BMS-" and len(nums[i])==15:
            ans.append(nums[i])
            flag=0
            break
    if flag:
        ans.append("UNK")
    
    #提取Subject No.受试者编号(patient id)
    flag=1
    for i in range(len(nums)):
        if nums[i]=='Patient' and nums[i+1]=='ID:':
            ans.append(nums[i+2])
            flag=0
            break
    if flag:
        ans.append("UNK")
        
    #提取性别
    flag=1
    for i in range(len(nums)):
        gender=nums[i].upper()
        if gender=="MALE":
            ans.append("M 男")
            flag=0
            break
        if gender=="FEMALE":
            ans.append("F 女")
            flag=0
            break
    if flag:
        ans.append("UNK")
        
    #提取年龄
    flag=1
    for i in range(10,len(nums)):
        if nums[i]=='ADVERSE' and nums[i+1]=='REACTION':
            for j in range(i-3,i):
                if nums[j].isupper():
                    if nums[j+1].isdigit():
                        ans.append(nums[j+1])
                        flag=0
                    elif nums[j+1]=="Unk":
                        ans.append("Unk")
                        flag=0
                    elif nums[j+1]=="Unknow":
                        ans.append("Unknow")
                        flag=0
            break
    if flag:
        ans.append("UNK")
        
    #提取国家 todo
    for i in range(len(nums)):
        if nums[i]=="Year" and nums[i+1]!="Day":
            c1=nums[i+1].upper()
            c2=c1+" "+nums[i+2].upper()
            c3=c2+" "+nums[i+3].upper()
            
            if c1 in CountryDict:
                ans.append(c1+" "+CountryDict[c1])
            elif c2 in CountryDict:
                ans.append(c2+" "+CountryDict[c2])
            elif c3 in CountryDict:
                ans.append(c3+" "+CountryDict[c3])
            else:
                ans.append("1")
            s=ans[-1]
            for i in range(len(s)):
                e=s[i]
                if '\u4e00' <= e <= '\u9fa5':
                    ans[-1]=s[:i]+"_"+s[i:]
                    break
                    
            #ans[-1]=ans[-1].replace('_','\r\n\t')
            break            
    #提取SAE TermSAE名称 todo
    """
    for i in range(len(nums)):
        if nums[i]=="Reportable" and nums[i+1]=="SAE:":
            j=i+2
            while nums[j][-1]!=")":
                j+=1
            ans.append(nums[i+2:j])
            break
    """
    ans.append(1)
    #提取发生时间 todo 
    flag=1
    for i in range(len(nums)):
        if nums[i]=='REACTION' and nums[i+1]=='ONSET':
            j=i+5
            if nums[j].isdigit() and (not nums[j+1].isdigit()) and nums[j+2].isdigit():
                ans.append(nums[j]+"-"+nums[j+1]+"-"+nums[j+2])
                flag=0
            elif nums[i+1]=="Unk":
                ans.append(nums[i-1])
                flag=0
    if flag:
        ans.append("UNK")
    #提取ManufacturerReceipt Date获知时间
    for i in range(len(nums)):
        if nums[i]=='BY' and nums[i+1]=='MANUFACTURER':
            ans.append(nums[i+2]) 
            print(nums[i-30:i+5])
    #提取报告时间 
    for i in range(len(nums)):
        if nums[i]=='THIS' and nums[i+1]=='REPORT':
            ans.append(nums[i+2]) 
    #提取Type of Safety Report随访类型
    for i in range(len(nums)):
        if nums[i]=="FOLLOWUP:":
            if len(nums[i+1])<3 and nums[i+1].isdigit():
                ans.append("Follow up "+nums[i+1]+" 随访- "+nums[i+1])
            else:
                ans.append("Initial 首次")    
    #提取Severity严重程度
    ans.append("1")
    #提取Death是否致死
    ans.append("1")
    #提取Action taked采取措施
    ans.append("1")
    #提取Outcome事件结局
    ans.append("1")
    #提取Nivo Causality(Investigator)Nivo 相关性(研究者)
    ans.append("1")
    #提取Nivo Causality(BMS)Nivo 相关性(BMS)
    ans.append("1")
    return ans

def f(nums):
    t=[0] * 5
    """
    for i in range(len(nums)):
    
        if nums[i]=='Year' and nums[i+1]!='Day':
            if nums[i+2].isdigit():
            
                #t.append([i+2,i+7])
                t=t+[i+2,i+7]+list(range(i+9,i+12))
                break
            else:
                #t.append([i+3,i+8])
                t=t+[i+3,i+8]+list(range(i+10,i+13))
                break
    """
    for i in range(len(nums)):
        if len(nums[i])>=4 and (nums[i][len(nums[i])-4:len(nums[i])]=='male' or nums[i][len(nums[i])-4:len(nums[i])]=='Male'):
            if nums[i-1]=='Unk':
                t[0]=i-1
                t[1]=i
                t[2]=i+1
                t[3]=i+2
                t[4]=i+3           
                break
            elif nums[i-1]=='Years':
                t[0]=i-5
                t[1]=i
                t[2]=i+1
                t[3]=i+2
                t[4]=i+3
                break
                
    for i in range(len(nums)):
        if nums[i]=='kg':
            t[2]=i+1
            t[3]=i+2
            t[4]=i+3
            break
    for i in range(len(nums)):
        if nums[i]=='PATIENT' and nums[i+1]=='DIED':
            if nums[i-1]=='Unk':
                t[2]=i-1
                t[3]=i-1
                t[4]=i-1
                break
    if not nums[t[0]].isdigit():
        if nums[t[0]]!='Unk':
            t[0]-=1
    for i in range(len(nums)):
        if nums[i]=='Patient' and nums[i+1]=='ID:':
            t.append(i+2)
            
            break
    for i in range(len(nums)):
    
        if nums[i]=='BY' and nums[i+1]=='MANUFACTURER':
            if nums[i+4]=='.':
                t.append(i+5)
            else:
            
                t.append(i+4)
            
    for i in range(len(nums)):
    
        if nums[i]=='REPORT' and nums[i+1]=='TYPE':
            if nums[i+2]=='.':
                t.append(i+3)
            else:
                t.append(i+2)           
    ans=[nums[t[5]],nums[t[0]],nums[t[1]],"-".join([nums[t[2]],nums[t[3]],nums[t[4]]]),nums[t[6]],nums[t[7]]]
    for i in range(len(nums)):
        if nums[i]=="FOLLOWUP:":
            if len(str(nums[i+1]))<3:
                ans.append("FOLLOWUP:"+nums[i+1])
            else:
                ans.append("Initial")
            
    #print(nums)
    for i in range(len(nums)):
        if nums[i][:4]=="BMS-" and len(nums[i])==15:
            ans.append(nums[i])
            break
    
    #if not ans[1].isdigit():
        #print(ans,nums)
    return ans 


