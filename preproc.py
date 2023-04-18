import os, glob, shutil, pickle

class Track():
    def __init__(self, DK):
        self.DK = DK
        self.rawFolder = f"/Volumes/2022_14TB/KatagiriProject/KP_cds/{self.DK}/"
        self.rawFiles = sorted(glob.glob(self.rawFolder+'*.aiff'))
        self.commands = [self.mergeRaw()]
        self.commands.append(self.convertToMP3())

    def returnLogs(self):
        return self.commands, self.rawFiles

    def mergeRaw(self):
        self.mergedRawFile = f"/Volumes/2022_14TB/KatagiriProject/KP_mergedDK/{self.DK}.aiff"
        with open('concat.txt','w') as fid:
            for file in self.rawFiles:
                fid.write(f"file '{file}'\n")
        #could this use python ffmpeg in the future?
        mycmd = f"/usr/local/bin/ffmpeg -y -f concat -safe 0 -i concat.txt -c copy {self.mergedRawFile}"
        #os.system(mycmd)
        print(mycmd)
        return mycmd


    def convertToMP3(self):
        self.mp3File = f"/Volumes/2022_14TB/KatagiriProject/KP_mp3/{self.DK}.mp3"
        mycmd = f"/usr/local/bin/ffmpeg -y -i {self.mergedRawFile} -f mp3 -acodec libmp3lame -b:a 24k -ar 24000 {self.mp3File}"
        #os.system(mycmd)
        print(mycmd)
        return mycmd

class TrackMultiCD(Track):
    def __init__(self, DKraws):
        self.DKraws = DKraws
        self.rawFiles = []
        for DKraw in self.DKraws:
            self.rawFiles.extend(sorted(glob.glob( f"/Volumes/2022_14TB/KatagiriProject/KP_cds/{DKraw}/*.aiff")) )
        self.DK = self.DKraws[0].split('-')[0]
        self.commands = [self.mergeRaw()]
        self.commands.append(self.convertToMP3())


def main():
    cdList = sorted(glob.glob(f"/Volumes/2022_14TB/KatagiriProject/KP_cds/DK*"))
    allcmds={}
    #pre-sort
    singleCDlist=[]
    multiCDdict={}
    for fullcd in cdList:
        cd = fullcd.split('/')[-1]
        if '-' in cd:
            DK, nums = cd.split('-')
            if DK in multiCDdict.keys():
                multiCDdict[DK].append(cd)
            else:
                multiCDdict[DK]=[cd]
        else:
            singleCDlist.append(cd)
    for cd in singleCDlist:
        track = Track(cd)
        allcmds[cd] = track.returnLogs()
    for DK in sorted(multiCDdict.keys()):
        track = TrackMultiCD(multiCDdict[DK])
        allcmds[DK] = track.returnLogs()
    with open('processSteps.pkl', 'wb') as fid:
        pickle.dump(allcmds, fid)
if __name__ == '__main__':
    main()