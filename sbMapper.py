import xlrd
import sys

def readExcelFile(file):
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_index(0)
    scoreListDict = {}
    for row in range(sheet.nrows-1):
        if sheet.cell_type(1+row,1)!=0:
            for i in range(int(sheet.cell(1+row,0).value)):
                main.append(sheet.cell(1+row,1).value)
        if sheet.cell_type(1+row,3)!=0:
            for i in range(int(sheet.cell(1+row,2).value)):
                side.append(sheet.cell(1+row,3).value)
    cell = sheet.cell(2,5)
    secEnd = 2;
    while(cell.value!="Deck"):
        secEnd+=1
        cell = sheet.cell(secEnd+1,5)
    for j in range(secEnd):
        r1 = 2+j
        r2 = secEnd+j+2
        rating = {}
        if sheet.cell(2+j,5).ctype == 0:
            break
        for col in range(7,sheet.ncols-1):
            if sheet.cell(1,col).ctype==1 and sheet.cell(r1,col).ctype!=0:
                rating.update({sheet.cell(1,col).value:sheet.cell(r1,col).value})
            if sheet.cell(secEnd+1,col).ctype==1 and sheet.cell(r2,col).ctype!=0:
                rating.update({sheet.cell(secEnd+1,col).value:sheet.cell(r2,col).value})
        if len(rating)>0:
            scoreListDict.update({sheet.cell(r1,5).value:rating})
            teirList.update({sheet.cell(r1,5).value:sheet.cell(r1,6).value})

    writelist(main,side)
    label = list(set(main)|set(side))
    return(scoreListDict)

def writelist(main,side):
    maindeck = {}
    sideboard = {}
    for card in basefile.readlines():
        count,card = card.rstrip().split(' ',1)
        maindeck[card] = int(count)
    for card in main:
        card = card.rstrip('1234567890')
        if card in maindeck:
            maindeck[card] += 1
        else:
            maindeck[card] = 1
    for card in side:
        card = card.rstrip('1234567890')
        if card in sideboard:
            sideboard[card] += 1
        else:
            sideboard[card] = 1
    for card,count in sorted(maindeck.items(),key=lambda x: x[1],reverse=True):
        outfile.write(str(count)+' '+card+'\n')
    outfile.write('Sideboard\n')
    for card,count in sorted(sideboard.items(),key=lambda x: x[1],reverse=True):
        outfile.write(str(count)+' '+card+'\n')
    outfile.write('\n')


def evaluate(name,scoreList):
    for card in scoreList:
        if card in main:
            scoreList.update({card:scoreList[card]+0.2})
    cards = sorted(scoreList, key=scoreList.get)
    out = 0
    valOut = 0
    valIn = 0
    outfile.write('~'+name+'\n')
    inDict = {}
    outDict = {}
    lines2write = []
    for card in cards:
        if card in side:
            copies = side.count(card)
            if out<15 and out+copies>15:
                copies = out+copies-15
                out = 15
            if out>=15:
                lines2write.append("+"+str(copies)+' '+str(card))
                currScore=round(scoreList[card])
                valIn += copies*currScore
                inDict.update({card:round(scoreList[card])})
            else:
                out+=copies
                outDict.update({card:round(scoreList[card])})
        if card in main:
            if out<15:
                outDict.update({card:round(scoreList[card])})
                copies = main.count(card)
                if copies+out>15:
                    copies=15-out
                lines2write.append("-"+str(copies)+' '+str(card))
                out+=copies
                if out >= 15:
                    currScore=(scoreList[card]-0.2)
                    valOut+=copies*currScore
            else:
                inDict.update({card:round(scoreList[card])})
    val = (valIn-valOut)
    firstChar = lines2write[0][0]
    lines2write.sort(reverse=True)
    l2wNoDup = [lines2write[0]]
    for i in range(1,len(lines2write)):
        if lines2write[i-1][3:].rstrip('1234567890') == lines2write[i][3:].rstrip('1234567890'):
            extras = lines2write[i-1]
            count,_ = extras[1:].split(' ',1)
            l = lines2write[i]
            print(l2wNoDup)
            l2wNoDup[-1] = l[0]+str(int(l[1])+int(count))+l[2:]
            print(l2wNoDup)
            print()
        else:
            l2wNoDup.append(lines2write[i])
    lines2write = l2wNoDup
    for line in lines2write:
        if line[0] != firstChar or line == lines2write[-1]:
            outfile.write(line.rstrip('1234567890')+'\n')
            firstChar = line[0]
        else:
            outfile.write(line+', ')
    maxOut = round(sorted(outDict.values(),reverse=True)[0])
    minIn = round(sorted(inDict.values())[0])
    #print(name,outDict,'\n',inDict)
    if maxOut == minIn:
        outfile.write('#Duplicate '+str(minIn)+'\'s: ')
        for card in cards:
            mcopies = main.count(card)
            scopies = side.count(card)
            if scopies>0 and round(scoreList[card]) == minIn and mcopies==0:
                outfile.write(" S"+str(scopies)+"_"+card)
            if mcopies>0 and round(scoreList[card]) == minIn:
                outfile.write(" M"+str(mcopies)+"_"+card)
        outfile.write('\n')
    if valIn == 0 and valOut == 0:
        outfile.write("No changes")
    else:
        outfile.write('#Score '+str(val))
    outfile.write('\n\n')
    return(val)

main = []
side =[] 
teirList = {}
outOf = 0
#scoreListDict = readTextfile("./list.txt")
basefile = open("baselist.txt",'r')
if len(sys.argv)>2:
    outfile = open(sys.argv[2],"w")
else:
    outfile = open("pre_map.txt","w")
if len(sys.argv)>1:
    scoreListDict = readExcelFile(sys.argv[1])
else:
    scoreListDict = readExcelFile("./currStorm.xlsx")
deckScore = 0
for name in sorted(scoreListDict.keys()):
    teirModifier = teirList[name]
    indivScore = evaluate(name,scoreListDict[name])*teirModifier
    deckScore+=indivScore
outval = deckScore/sum(teirList.values())
outfile.write("Total score:{: .1f}".format(outval))
