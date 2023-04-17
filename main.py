import itertools, pandas, docx, glob, os, shutil
from mutagen import id3

inDir = '/Volumes/2022_14TB/KatagiriProject/KP_cds/'
mergedWithinCDdir =  '/Volumes/2022_14TB/KatagiriProject/KP_mergedWithinCD/'
mergedDKdir =  '/Volumes/2022_14TB/KatagiriProject/KP_mergedDK/'
outDir = '/Volumes/2022_14TB/KatagiriProject/KP_mp3/'

infoDir = '/Users/agrant/codes/MZMC/katagiri/2023/'
inFile = 'KATAGIRI MASTER CATALOG.docx'
#inFile = 'KATAGIRI MASTER CATALOGdebug.docx'
outFile = 'KATAGIRI MASTER CATALOG_ANG2.docx'


def pre_process_audio():
    cdList = sorted(glob.glob(f"{inDir}DK*"))
    multiCDlist=[]
    DKlist=[]
    for dir in cdList:
        #merge within dir
        DK = dir.split('/')[-1]
        concatOrigCDtracks(dir, DK)
        if '-' in DK:
            multiCDlist.append(DK.split('-')[0])
            DKlist.append(DK.split('-')[0])
        else:
            DKlist.append(DK)
    multiCDlist = set(multiCDlist)
    DKlist = set(DKlist)
    for DK in DKlist:
        if DK in multiCDlist:
            mergeCDs(DK)
        else:
            copyToMergedDK(DK)
        convertAIFFtoMP3(DK)

def copyToMergedDK(DK):
    inFile = f"{mergedWithinCDdir}{DK}.aiff"
    shutil.copy2(inFile, f"{mergedDKdir}{DK}.aiff")


def mergeCDs(DK):
    #merge multiple cds
    fileList=sorted(glob.glob(f"{mergedWithinCDdir}{DK}*.aiff"))
    thisOutDir = f"{mergedDKdir}"
    concatTracks(fileList, thisOutDir, DK)
def convertAIFFtoMP3(DK):
    inFile = f"{mergedDKdir}{DK}.aiff"
    outFile = f"{outDir}{DK}.mp3"
    mycmd = f"/usr/local/bin/ffmpeg -y -i {inFile} -f mp3 -acodec libmp3lame -b:a 24k -ar 24000 {outFile}"
    os.system(mycmd)

def concatOrigCDtracks(dir, DK):
    fileList = sorted(glob.glob(f"{dir}/*.aiff"))
    thisOutDir = f"{mergedWithinCDdir}"
    concatTracks(fileList, thisOutDir,DK)

def concatTracks(fileList, thisoutdir, DK):
    if os.path.exists(thisoutdir) is False:
        os.mkdir(thisoutdir)
    if len(fileList)==1:
        #single file--copy aiff to dstDir
        shutil.copy2(fileList[0],f"{thisoutdir}/{DK}.aiff")
    elif len(fileList)>1:
        #multiple files--concat into dstDir
        with open('concat.txt','w') as fid:
            for file in fileList:
                fid.write(f"file '{file}'\n")
        mycmd = f"/usr/local/bin/ffmpeg -y -f concat -safe 0 -i concat.txt -c copy {thisoutdir}/{DK}.aiff"
        os.system(mycmd)


def loadDoc(file):
    doc = docx.Document(file)
    tables = doc.tables
    table = tables[0]
    data = [[cell.text for cell in row.cells] for row in table.rows]
    oldhead = data[0]
    data = data[1:]
    return data, oldhead


class OutDoc:
    def __init__(self, file, outCols):
        self.document = docx.Document()
        self.file = file
        self.outCols = outCols
        self.section = self.document.sections[0]
        self.section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        self.section.page_width = docx.shared.Inches(17)
        self.section.page_height = docx.shared.Inches(11)
        self.table = self.document.add_table(rows=1, cols=len(outCols), style="Table Grid")
        headerCells = self.table.rows[0].cells
        for iC, col in enumerate(outCols):
            headerCells[iC].text = col
            # print(iC, col)

    def addRow(self, trackInfo):
        rowCells = self.table.add_row().cells
        for iK, key in enumerate(self.outCols):
            rowCells[iK].text = str(trackInfo[key])
            # print(f'{key}: {str(trackInfo[key])} ..... {rowCells[iK].text}')

    def write(self):
        self.document.save(self.file)


class TableRow():
    def __init__(self, row, rowtype, oldhead, newhead):
        self.oldhead = oldhead
        self.newhead = newhead
        self.tableRowOutput = []
        if rowtype == 'tracks':
            # create trackInfo but don't touch audio yet
            self.df = row
            self.parseTracks()
        elif rowtype == 'info':
            self.row = row
            self.dummyTrack()
            # create a dummytrack formatted to append to the output table

    def parseTracks(self):
        self.tracks = []
        if self.df.shape[0] == 1:
            ind = 0
            # single track in this tablerow
            trackTitleRaw = self.df.loc[ind, 'Title']
            if trackTitleRaw is None:
                trackTitle = self.df.loc[0, 'Title']
                trackComment = ''
            elif '(' in trackTitleRaw:
                titleParts = trackTitleRaw.split('(')
                trackTitle = titleParts[0]
                trackComment = '(' + ' '.join(titleParts[1:])
            else:
                trackTitle = trackTitleRaw
                trackComment = ''
            if ',' in self.df.loc[ind, 'Digital ID']:
                digID = self.df.loc[ind, 'Digital ID'].split('-')[0]
            else:
                digID=self.df.loc[ind, 'Digital ID']
            trackInfo = {'Series Title': trackTitle,
                         'Series Comment': trackComment,
                         'Digital ID': digID,
                         'Title': trackTitle,
                         'Title Comment': trackComment,
                         'Track Number': 1,
                         'Total Tracks': 1,
                         'Date': self.df.loc[ind, 'Date'],
                         'Digital Formats': self.df.loc[ind, 'Digital Formats'],
                         'Online': self.df.loc[ind, 'Online'],
                         'Flag': 0}
            for col in ['Occasion', 'C', 'T', 'P']:
                if (self.df.loc[ind, col] is None) or (self.df.loc[ind, col] == ''):
                    trackInfo[col] = self.df.loc[0, col]
                else:
                    trackInfo[col] = self.df.loc[ind, col]
            self.tracks.append(trackInfo)
            # print(trackInfo)
        else:
            # multiple rows
            comments = []
            seriestitle = self.df.loc[0, 'Title']
            skipInd = []

            for ind in self.df.index:
                if self.df.loc[ind, 'Digital ID'] is None:
                    if self.df.loc[ind, 'Title'] is not None:
                        comments.append(self.df.loc[ind, 'Title'])
                    skipInd.append(ind)
                elif self.df.loc[ind, 'Digital ID'] == '':
                    if self.df.loc[ind, 'Title'] is not None:
                        comments.append(self.df.loc[ind, 'Title'])
                    skipInd.append(ind)
            trackNum = 1
            for ind in self.df.index:
                if ind not in skipInd:
                    title = self.df.loc[0, 'Title']
                    # print(comments)
                    if len(comments) > 0:
                        thiscomment = ' '.join(comments)
                    else:
                        thiscomment = ''
                    if ',' in self.df.loc[ind, 'Digital ID']:
                        digID = self.df.loc[ind, 'Digital ID'].split('-')[0]
                    else:
                        digID = self.df.loc[ind, 'Digital ID']
                    trackInfo = {'Series Title': seriestitle,
                                 'Series Comment': thiscomment,
                                 'Digital ID': digID,
                                 'Title': title,
                                 'Title Comment': '',
                                 'Track Number': trackNum,
                                 'Total Tracks': self.df['Digital ID'].str.count('DK').sum(),
                                 'Date': self.df.loc[ind, 'Date'],
                                 'Digital Formats': self.df.loc[ind, 'Digital Formats'],
                                 'Online': self.df.loc[ind, 'Online'],
                                 'Flag': 0}
                    for col in ['Occasion', 'C', 'T', 'P']:
                        if self.df.loc[0, col] is not None:
                            trackInfo[col] = self.df.loc[0, col]
                    self.tracks.append(trackInfo)
                    # print(trackInfo)
                    trackNum += 1

    def dummyTrack(self):
        # parse the raw list into a dict
        trackInfo = {}
        for head in self.newhead:
            if head == 'Title':
                trackInfo[head] = self.row[2]
            else:
                trackInfo[head] = ''
        self.tracks = [trackInfo]
        # print(trackInfo)

    def writeRows(self, doc):
        for track in self.tracks:
            # loop through trackinfo and add rows to output table
            doc.addRow(track)

    def processTracks(self):
        # loop through trackinfo and touch audio by creating the objects
        for thistrackInfo in self.tracks:
            track = Track(thistrackInfo)
            trackOutput = track.tag()
            self.tableRowOutput.extend(trackOutput)
        return self.tableRowOutput



class Track:
    def __init__(self, trackInfo):
        self.trackInfo = trackInfo
        #self.debugMode = True
        self.debugMode = False
        self.trackOutput=[]
        self.outMP3 = f"{outDir}{self.trackInfo['Digital ID']}.mp3"

    def tag(self):
        # tag mp3
        if self.debugMode == True:
            print(f"tagging {self.outMP3}")
        else:
            tags = id3.ID3(self.outMP3)
            tags['TIT2'] = id3.TIT2(encoding=3, text=self.trackInfo['Title'])  # title of track
            tags['TRCK'] = id3.TRCK(encoding=3,
                                    text=f"{self.trackInfo['Track Number'] + 1}/{self.trackInfo['Total Tracks']}")  # track number
            tags['TPE1'] = id3.TPE1(encoding=3, text='Dainin Katagiri Roshi')  # artist
            tags['TCON'] = id3.TCON(encoding=3, text='Speech')  # genre
            year = int(self.trackInfo['Date'].split('/')[2]) + 1900
            tags['TDRC'] = id3.TDRC(encoding=3, text=f'{year}')  # year
            tags['TALB'] = id3.TALB(encoding=3, text=f"{self.trackInfo['Series Title']}")  # album
            tags['COMM'] = id3.COMM(encoding=3, lang=u'eng', desc='desc', text=self.trackInfo['Series Comment'])  # comments
            tags['TCOM'] = id3.TCOM(encoding=3, text='Dainin Katagiri Roshi')  # composer

            picFID = open('/Volumes/2022_14TB/KatagiriProject/KP_cds/pix.jpg', 'rb')
            # https://stackoverflow.com/questions/47346399/how-do-i-add-cover-image-to-a-mp3-file-using-mutagen-in-python
            tags['APIC:'] = id3.APIC(
                encoding=3,  # 3 is for utf-8
                mime='image/png',  # image/jpeg or image/png
                type=3,  # 3 is for the cover image
                desc=u'Cover',
                data=picFID.read()
            )  # picture!
            picFID.close()
            tags.save(self.outMP3)
        self.trackOutput.append(f"tagging {self.outMP3}")



def main():
    preProc = pre_process_audio()

    data, oldhead = loadDoc(infoDir + inFile)
    newhead = ['Series Title', 'Series Comment',
               'Digital ID', 'Title', 'Title Comment', 'Track Number',
               'Total Tracks', 'Date', 'Occasion', 'C', 'T', 'P', 'Digital Formats', 'Online', 'Flag']
    outdoc = OutDoc(infoDir + outFile, newhead)
    with open('outNotes.txt','w') as fidOutNotes:
        for row in data:
            if 'DK' in row[0]:
                # talk with tracks--go process them
                rowlist = []  # start with a list of lists
                for col in row:
                    if '\n' in col:
                        rowlist.append(col.split('\n'))
                    else:
                        rowlist.append([col])
                # OOOOOOOHHHH https://stackoverflow.com/questions/46431660/create-a-pandas-dataframe-from-a-nested-lists-of-unequal-lengths
                # magically expand to an array in a dataframe
                df = pandas.DataFrame((_ for _ in itertools.zip_longest(*rowlist)), columns=oldhead)
                tablerow = TableRow(df, 'tracks', oldhead, newhead)
                tablerow.writeRows(outdoc)
                tableRowOutput = tablerow.processTracks()
                fidOutNotes.writelines('\n'.join(tableRowOutput))

            else:
                # some other kind of entry--just pass it along to the output table
                tablerow = TableRow(row, 'info', oldhead, newhead)
                tablerow.writeRows(outdoc)

    outdoc.write()


if __name__ == '__main__':
    main()