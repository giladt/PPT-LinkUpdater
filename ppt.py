from win32com import client
from tkinter import Tk, Label, X, BOTH
from tkinter import filedialog, ttk

class OfficeApp:

    def __init__(self, AppIdentifier):
        self.SourceFolder = ''
        self.FileName = ''
        self.Presentation = None
        self.AppIdentifier = AppIdentifier

    def FullPath(self, strPath = ''):
        if strPath == '':
            _FullPath = r'{}\{}'.format(self.SourceFolder,self.FileName)
        else:
            _FullPath = strPath.replace('/','\\')
            self.FileName = _FullPath.split('\\')[-1]
            self.SourceFolder = _FullPath[0:int(_FullPath.rfind('\\'))]

        return _FullPath

    def Open(self):
        if self.AppIdentifier == "PowerPoint.Application":
            self.Application = client.Dispatch(self.AppIdentifier)
            bIsOpen = False
            for prs in self.Application.Presentations:
                if prs.name == self.FileName:
                    bIsOpen = True

            if not bIsOpen:
                #print("Opening file: " + self.FileName)
                self.Presentation = self.Application.Presentations.Open(FileName = self.FullPath(), WithWindow = 0)
            else:
                #print(self.FileName + " is already open.")
                self.Presentation = self.Application.Presentations(self.FileName)

            self.Application = None
            return bIsOpen

    def Close(self):
        if self.AppIdentifier == "PowerPoint.Application":
            self.Application = client.Dispatch(self.AppIdentifier)
            if self.Presentation != None:
                #print("Closing presentation " + self.FileName)
                self.FileName = ""
                self.Presentation.Close()
                self.Application.Quit()
                del self.Application
                #print("Closed PPT")

    def GetLinksFromSlides(self, posx,posy):
        if self.AppIdentifier == "PowerPoint.Application":
            self.FileName = filedialog.askopenfilename(
                title = "Select a PowerPoint file",
                filetypes = (("PowerPoint files", "*.pptx *.pptm"), ("All files", "*.*"))
            )
            #print("Getting links from file:", self.FileName)

            arrLinks=[]
            if self.FileName != '':
                self.progress = Tk()
                self.progress.geometry('250x100+' + str(posx - 125) + '+' + str(posy - 50))
                self.progress.title("Getting links")
                self.progress.iconbitmap(r'.\trendence_black.ico')
                
                Label(self.progress, text="Importing links").pack(fill=X)
                self.bar = ttk.Progressbar(self.progress, orient = 'horizontal', length=100, mode='determinate')
                self.bar.pack(fill=BOTH)
                self.FullPath(self.FileName)
                self.Open()

                #print("Gathering links")
                max=0
                for oSlide in self.Presentation.slides:
                    max=max+oSlide.shapes.count

                count=0
                self.bar['maximum']=max
                for oSlide in self.Presentation.slides:
                    for oShape in oSlide.shapes:
                        count = count + 1
                        self.bar['value'] = count
                        self.bar.update()

                        if oShape.type == 11:
                            arrLinks.append({
                                'shape': oShape,
                                'link': oShape.LinkFormat.SourceFullName
                            })

                #print("Done")
                self.progress.destroy()
                #print("Progress window distroyed")
            return arrLinks

    def SetLinksInSlides(self, arrLinks, posx,posy):
        if self.AppIdentifier == "PowerPoint.Application":
            #print("Setting links in file:", self.FileName)

            if len(arrLinks) > 0:
                self.progress = Tk()
                self.progress.geometry('250x100+' + str(posx - 125) + '+' + str(posy - 50))
                self.progress.title("Replacing links")
                self.progress.iconbitmap(r'.\trendence_black.ico')
                
                Label(self.progress, text="Exporting links").pack(fill=X)
                self.bar = ttk.Progressbar(self.progress, orient = 'horizontal', length=100, mode='determinate')
                self.bar.pack(fill=BOTH)
                self.FullPath(self.FileName)

                max=len(arrLinks)
                self.bar['maximum']=max

                count=0
                for item in arrLinks:
                    item['shape'].LinkFormat.SourceFullName = item['link']
                    count = count + 1
                    self.bar['value'] = count
                    self.bar.update()

                self.FullPath(self.FileName)
                #print(self.FullPath)
                _bIsOpen = self.Open()
                #print(_bIsOpen)
                self.Presentation.Save()
                #print("Done")
                self.progress.destroy()
                #print("Progress window distroyed")
