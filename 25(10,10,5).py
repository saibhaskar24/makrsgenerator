import pandas as pd
dff = pd.ExcelFile('INTERNAL-SHEET.xlsx')
minn=[]
writer = pd.ExcelWriter("out.xlsx")
for sheet in dff.sheet_names:
    print(sheet)
    df = pd.read_excel('INTERNAL-SHEET.xlsx',sheet)
    regno = df['REG']
    internal=df['INTERNAL']
    ll=[]
    import random
    res = {"REG":[],"Assignment":[],"Proj":[],"Mid":[],"Total":[]}
    for I in range(0,len(df)):
        res["REG"].append(regno[I])
        res["Total"].append(internal[I])
        total = internal[I]
        assig = 5
        if total > 10 and total < 20:
            assig = random.randint(3,4)
        elif total <= 10 and total >= 5:
            assig = 2
        elif total < 5 and total >= 2:
            assig = 1
        elif total ==0:
            assig = 0
        total -= assig
        proj = random.randint(7,10)
        if total < 10 and total > 7:
            proj = random.randint(6,8)
        elif total <=7 and total > 5:
            proj = random.randint(4,6)
        elif total <=6 and total > 4:
            proj = random.randint(3,5)
        elif total <=4 and total >= 2:
            proj = random.randint(1,2)
        elif total ==0:
            proj = 0
        total -= proj
        if total < 0 or assig + proj + total != internal[I] and total > proj:
            print("Total",total,"Internal", internal[I])
        res["Assignment"].append(assig)
        res["Mid"].append(total)
        res["Proj"].append(proj)        
    resdf = pd.DataFrame(res)
    resdf.to_excel(writer,sheet)
writer.save()


            
    

            


