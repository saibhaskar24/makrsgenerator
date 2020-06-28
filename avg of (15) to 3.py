import pandas as pd
import math
import numpy as np
dff = pd.ExcelFile('2014 BATCH INTERNALS.xlsx')
minn=[]
writer = pd.ExcelWriter("out2014.xlsx")
import random

def div15(t):
    if t < 3:
        return [t,t,t]
    if t==15:
        return [t,t,t]
    if t==14:
        h = random.randint(0,5)
        if h ==0:
            return [t,t,t]
        if h == 1:
            return [t-1,t,t+1]
        if h == 2:
            return [t,t,t-1]
        else:
            return [t,t-1,t]

    h = random.randint(-1,1)
    d = []
    dd = []
    l = [0]*3
    p = random.randint(0,2)
    d.append(p)
    l[p] = t+h
    dd.append(h)

    while p in d:
        p = random.randint(0,2)
        while h in dd:
            h = random.randint(-1,1)
    d.append(p)
    l[p] = t+h

    if 1 in d and 2 in d:
        l[0] = 3*t - sum(l)
    if 1 in d and 0 in d:
        l[2] = 3*t - sum(l)
    if 0 in d and 2 in d:
        l[1] = 3*t - sum(l)

    


    return l





for sheet in dff.sheet_names:
    print(sheet)
    df = pd.read_excel('2014 BATCH INTERNALS.xlsx',sheet)
    regno = df['s.no']
    sub1=df['sub1']
    sub2=df['sub2']
    res = {"s.no":[],"sub1":[],"q1":[],"q2":[],"q3":[],"sub2":[],"q4":[],"q5":[],"q6":[]}
    for I in range(0,len(df)):
        res["s.no"].append(regno[I])
        res["sub1"].append(sub1[I])
        if str(sub1[I])!='nan' :
            try:
                k = div15(sub1[I])
                if math.ceil(sum(k)/3) != sub1[I]:
                    print("Mistake in 1 : ",k, sub1[I])
                for ikk in k:
                    if ikk<0 or ikk>15:
                        print("Error 1: ",k, sub1[I])
                        break
                res["q1"].append(k[0])
                res["q2"].append(k[1])
                res["q3"].append(k[2])
            except:
                res["q1"].append('A')
                res["q2"].append('A')
                res["q3"].append('A')
        else:
            # print(sub1[I])
            res["q1"].append('')
            res["q2"].append('')
            res["q3"].append('')

        res["sub2"].append(sub2[I])
        if str(sub2[I]) != 'nan':
            try:
                k = div15(sub2[I])
                if math.ceil(sum(k)/3) != sub2[I]:
                    print("Mistake in 2 : ",k , sub2[I])
                for ikk in k:
                    if ikk<0 or ikk>15:
                        print("Error 2: ",k, sub2[I])
                        break
                res["q4"].append(k[0])
                res["q5"].append(k[1])
                res["q6"].append(k[2])
            except:
                res["q4"].append('A')
                res["q5"].append('A')
                res["q6"].append('A')

        else:
            # print(sub2[I])
            res["q4"].append('')
            res["q5"].append('')
            res["q6"].append('')
        



    resdf = pd.DataFrame(res)
    resdf.to_excel(writer,sheet)
writer.save()

