import xlrd
import sys
import argparse

def readExcelFile(file):
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_index(0)
    scoreListDict = {}
    teirList = {}
    main = []
    side = []
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

    #label = list(set(main)|set(side))
    return(scoreListDict,main,side,teirList)

def writelist(main,side,basefn,outfile):
    with open(basefn,'r') as basefile:
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

def evaluate(name,scoreList,main,side,outfile,max_score):
    for card in scoreList:
        if card in main:
            scoreList.update({card:scoreList[card]+0.2})
    cards = sorted(scoreList, key=scoreList.get)
    out = 0
    outfile.write('~'+name+'\n')
    lines2write = []
    bringIn = {}
    takeOut = {}
    extras = []

    maxOut = 0
    minIn = max_score
    rem = 15
    for card in cards:
        cardScore = round(scoreList[card])
        if '!' in card:
            card = card.rstrip('!')
        mc = main.count(card)
        sc = side.count(card)
        if mc == 0 and sc == 0:
            continue
        if mc + sc <= rem:
            maxOut = cardScore
            rem -= mc + sc
            if mc > 0:
                if card in takeOut:
                    takeOut[card] += mc
                else:
                    takeOut[card] = mc
        elif rem == 0:
            minIn = min(minIn,cardScore)
            if sc > 0:
                bringIn[card] = sc
        else:
            if sc <= rem:
                rem -= sc
            else:
                bringIn[card] = sc - rem
                minIn = min(minIn,cardScore)
                if rem < sc:
                    maxOut = cardScore
                rem = 0
            if rem > 0:
                if mc > 0:
                    if card in takeOut:
                        takeOut[card] += rem
                    else:
                        takeOut[card] = rem
                rem = 0
    max_effic = len(main)*max_score
    totval = 0
    for card in main:
        totval += min(max_score,round(scoreList[card]))
    preScore = str(int((100*totval)/max_effic))

    if len(bringIn) == 0 and len(takeOut) == 0:
        outfile.write('#No changes, efficacy '+preScore+'%\n\n')
        return(totval)

    valOut = 0
    valIn = 0
    for card,count in sorted(takeOut.items(),key=lambda x: str(x[1])+x[0],reverse = True):
        outfile.write('-'+str(count)+' '+card+' ')
        valOut += int(count)*min(max_score,round(scoreList[card]))
    outfile.write('\n')
    for card,count in sorted(bringIn.items(),key=lambda x: str(x[1])+x[0],reverse = True):
        outfile.write('+'+str(count)+' '+card+' ')
        valIn += int(count)*min(max_score,round(scoreList[card]))

    outfile.write('\n')
    val = valIn - valOut
    if maxOut == minIn:
        outfile.write('#Chose between multiple '+str(minIn)+'\'s: ')
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

def eval_list(scoreListDict,main,side,teirList,outfile,max_score):
    deckScore = 0
    scores = {}
    for name in sorted(scoreListDict.keys()):
        teirModifier = teirList[name]
        indivScore = evaluate(name,scoreListDict[name].copy(),main,side,outfile,max_score)
        scores[name] = indivScore
        indivScore *= teirModifier
        deckScore+=indivScore
    outval = str(deckScore/sum(teirList.values()))
    outfile.write("#Average postboard efficacy:"+outval+"%\n")
    for d,s in sorted(scores.items(),key=lambda x: x[1]):
        print(d,s)
    print("Average postboard efficacy:"+outval+"%\n")

def argparser():
    parser = argparse.ArgumentParser()
    parser.add_argument('path',nargs='?')
    parser.add_argument('-o',nargs='?', default="pre_map.txt")
    parser.add_argument('-m',nargs='?', default=6)
    args = parser.parse_args()
    fname = vars(args)['path']
    basefile = "baselist.txt"
    outfile = open(vars(args)['o'],"w")
    max_score = int(vars(args)['m'])
    scoreListDict,main,side,teirList = readExcelFile(fname)
    writelist(main,side,basefile,outfile)
    eval_list(scoreListDict,main,side,teirList,outfile,max_score)

if __name__ == "__main__":
    argparser()
