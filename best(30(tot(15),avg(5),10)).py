import pandas as pd
dff = pd.ExcelFile('ss.xlsx')
minn=[]
writer = pd.ExcelWriter("2015-19 INTERNAL final year.xlsx")
for sheet in dff.sheet_names:
    print(sheet)
    if "Table" in sheet or "LAB" in sheet or "ROUGH" in sheet:
        continue
    df = pd.read_excel('ss.xlsx',sheet)
    m1 = df['m1']
    m2 = df['m2']
    l=[]
    l.append(m1)
    l.append(m2)
    regno = df['htno']
    sq1=df['sq1']
    assign1=df['assign1']
    sum1=df['sum1']
    sum2=df['sum2']
    best=df['best']
    sq2=df['sq2']
    assign2=df['assign2']
    ll=[]
    ll.append(assign1)
    ll.append(assign2)
    
    import random
    res = {"Reg no":[],"Q11":[],"Q12":[],"Q13":[],"mid1":[],"sq1":[],"A11":[],"A12":[],"A13":[],"assign1":[],"sum1":[],"Buffer":[],"Q21":[],"Q22":[],"Q23":[],"mid2":[],"sq2":[],"A21":[],"A22":[],"A23":[],"assign2":[],"sum2":[],"best":[]}
    for I in range(0,len(m1)):
        res["Reg no"].append(regno[I])
        res["Buffer"].append(" ")
        res["sq1"].append(sq1[I])
        res["sq2"].append(sq2[I])
        res["assign1"].append(assign1[I])
        res["assign2"].append(assign2[I])
        res["sum1"].append(sum1[I])
        res["sum2"].append(sum2[I])
        res["best"].append(best[I])
        for II in range(2):
            i = int(l[II][I])
            r = i//3
            z1 = random.randint(r,5)
            z2 = random.randint(0,r)
            z3 = i - z1 - z2
            if z3<0:
                    z1 = z1+z3
                    z3 = 0
            if z3>5:
                    z2 = z2 + (z3-5)
                    z3 = 5
            if z3 < 0 or z3>5 or z2 < 0 or z2>5 or z1 < 0 or z1>5 or z1+z2+z3!=i:
                minn.append(z3)
                print("Unhandled for ",i," values are :",z1,z2,z3)
            rt = random.randint(0,1)
            if rt ==0:
                z1,z2 = z2,z1
            res["Q"+str(II+1)+"1"].append(z1)
            res["Q"+str(II+1)+"3"].append(z2)
            res["Q"+str(II+1)+"2"].append(z3)
            j = int(ll[II][I])
            rt = random.randint(0,4)
            if rt == 0:
                res["A"+str(II+1)+"1"].append(j)
                res["A"+str(II+1)+"3"].append(j)
                res["A"+str(II+1)+"2"].append(j)
            elif rt == 1:
                res["A"+str(II+1)+"1"].append(j-1)
                res["A"+str(II+1)+"3"].append(j)
                res["A"+str(II+1)+"2"].append(j)
            elif rt == 2:
                res["A"+str(II+1)+"1"].append(j)
                res["A"+str(II+1)+"3"].append(j-1)
                res["A"+str(II+1)+"2"].append(j)
            elif rt == 3:
                res["A"+str(II+1)+"1"].append(j)
                res["A"+str(II+1)+"3"].append(j)
                res["A"+str(II+1)+"2"].append(j-1)
            else:
                res["A"+str(II+1)+"1"].append(j)
                res["A"+str(II+1)+"3"].append(j-1)
                res["A"+str(II+1)+"2"].append(j-1)        
            res["mid"+str(II+1)].append(i)        
    resdf = pd.DataFrame(res)
    resdf.to_excel(writer,sheet)
writer.save()
print(minn)


            
    

            


