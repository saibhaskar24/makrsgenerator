import pandas as pd
dff = pd.ExcelFile('2-2.xlsx')
print(dff.sheet_names)
writer = pd.ExcelWriter("res.xlsx")
for sheet in dff.sheet_names:
    df = pd.read_excel('2-2.xlsx',sheet)
    l = df['Internal']
    regno = df['Htno']
    #print(df)
    import random
    res = {"Reg no":[],"A1":[],"A2":[],"A3":[],"Quiz1":[],"Q1":[],"Q2":[],"Q3":[],"Total":[]}
    #l = [15,14,13,15,13]
    for I in range(0,len(l)):
        
        i = int(l[I])
        if i<6:
            x=4
        else:
            x=5
        rr = i-x
        rr = int(rr * 2/5)
        y = random.randint(rr//3,10)
        if i < 20 and y>5:
            y = random.randint(3,5)
        if i < 10 and y>3:
            y = random.randint(0,3)
        z = i - y - x
        if z > 15:
            y = y+(z-15)
            z = 15
        z = i - y - x
        if z <= 15:
            rr = z//3
            z1 = random.randint(rr,5)
            z2 = random.randint(rr,5)
            z3 = z - z1 - z2
            if z3 > 5:
                if z3 == 6:
                    ttt  = random.randint(0,1)
                    if ttt == 1:
                        z2 = 5
                    else:
                        z1 = 5
                    z3 = 5
                else:
                    print("Unhandled 15",z3)
            if z3 < 0:
                print("Unhandled neg")
            res["Reg no"].append(regno[I])
            res["Q1"].append(z1)
            res["Q2"].append(z2)
            res["Q3"].append(z3)
            ra = random.randint(0,2)
            if ra == 0:
                res["A1"].append(x-1)
                res["A3"].append(x)
                res["A2"].append(x)
            elif ra == 1:
                res["A2"].append(x-1)
                res["A1"].append(x)
                res["A3"].append(x)
            else:
                res["A3"].append(x-1)
                res["A1"].append(x)
                res["A2"].append(x)
            res["Quiz1"].append(y)
            res["Total"].append(i)
            print(x,y,z1,z2,z3,i)
        else:
            print("Unhandled")
            print(x,y,z,i)
        
        
    resdf = pd.DataFrame(res)
    resdf.to_excel(writer,sheet)
writer.save()
            
    

            
    
