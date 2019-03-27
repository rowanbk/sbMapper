import xlrd
import sys
import argparse

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
    #label = list(set(main)|set(side))
    return(scoreListDict)

def writelist(main,side):
    maindeck = {}
    sideboard = {}
    for card in basefile.readlines():
        count,card = card.rstrip().split(' ',1)
        maindeck[card] = int(count)
    for card in main:
        card = card.rstrip('!')
        if card in maindeck:
            maindeck[card] += 1
        else:
            maindeck[card] = 1
    for card in side:
        card = card.rstrip('!')
        if card in sideboard:
            sideboard[card] += 1
        else:
            sideboard[card] = 1
    for card,count in sorted(maindeck.items(),key=lambda x: str(x[1])+x[0],reverse=True):
        outfile.write(str(count)+' '+card+'\n')
    outfile.write('Sideboard\n')
    for card,count in sorted(sideboard.items(),key=lambda x: str(x[1])+x[0],reverse=True):
        outfile.write(str(count)+' '+card+'\n')
    outfile.write('\n')


def evaluate(name,scoreList):
    for card in scoreList:
        if card in main:
            scoreList.update({card:scoreList[card]+0.2})
    cards = sorted(scoreList, key=scoreList.get)
    out = 0
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
            else:
                inDict.update({card:round(scoreList[card])})
    max_effic = len(main)*MAX_SCORE
    totval = 0
    for card in main:
        totval += min(MAX_SCORE,round(scoreList[card]))
    preScore = str(int((100*totval)/max_effic))

    if len(lines2write) == 0:
        outfile.write('No changes, efficacy '+preScore+'%\n\n')
        return(totval)

    valOut = 0
    valIn = 0
    for count,card in [c[1:].split(' ',1) for c in lines2write if '+' in c]:
        valIn += int(count)*min(MAX_SCORE,round(scoreList[card]))
    for count,card in [c[1:].split(' ',1) for c in lines2write if '-' in c]:
        valOut += int(count)*min(MAX_SCORE,round(scoreList[card]))
    val = valIn - valOut

    firstChar = lines2write[0][0]
    lines2write.sort(reverse=True)
    l2wNoDup = [lines2write[0]]
    for i in range(1,len(lines2write)):
        if lines2write[i-1][3:].rstrip('!') == lines2write[i][3:].rstrip('!'):
            extras = lines2write[i-1]
            count,_ = extras[1:].split(' ',1)
            l = lines2write[i]
            l2wNoDup[-1] = l[0]+str(int(l[1])+int(count))+l[2:]
        else:
            l2wNoDup.append(lines2write[i])
    lines2write = l2wNoDup

    for line in lines2write:
        if line[0] != firstChar:
            outfile.write('\n')
            firstChar = line[0]
        outfile.write(line.rstrip('!')+' ')
        if line == lines2write[-1]:
            outfile.write('\n')
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
    posScore = str(int((100*(totval+val))/max_effic))
    outfile.write('#Efficacy '+posScore+'% improved from '+preScore+'% preboard\n')
    outfile.write('\n\n')
    return(int(posScore))

def inc(cc):
    for p in reversed(range(len(cc)-1)):
        if cc[p] < 4:
            cc[p] += 1
            return p,cc
    return None

def dec(cc,d):
    for p in range(d+1,len(cc)):
        if cc[p] > 0:
            cc[p] -= 1
            return p,cc
    return None

def eval_xlsx(fname):
    scoreListDict = readExcelFile(fname)
    deckScore = 0
    for name in sorted(scoreListDict.keys()):
        teirModifier = teirList[name]
        indivScore = evaluate(name,scoreListDict[name])*teirModifier
        deckScore+=indivScore
    outval = str(deckScore/sum(teirList.values()))
    outfile.write("Average postboard efficacy:"+outval+"%\n")
    print("Average postboard efficacy:"+outval+"%\n")

parser = argparse.ArgumentParser()
parser.add_argument('path',nargs='?')
parser.add_argument('-o',nargs='?', default="pre_map.txt")
parser.add_argument('-m',nargs='?', default=4)
args = parser.parse_args()

teirList = {}
outOf = 0
#scoreListDict = readTextfile("./list.txt")
fname = vars(args)['path']
basefile = open("baselist.txt",'r')
outfile = open(vars(args)['o'],"w")
MAX_SCORE = int(vars(args)['m'])
main = []
side = []

if __name__ == "__main__":
    eval_xlsx(fname)
