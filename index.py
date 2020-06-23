from tkinter import Tk, Label, Entry, LabelFrame, \
                    Frame, Button, Listbox, Scrollbar, \
                    LEFT, RIGHT, TOP, BOTTOM, BOTH, EXTENDED, FLAT, DISABLED, NORMAL, \
                    X, Y, W, E, N, S, NW, SE, NSEW, END, messagebox, ttk

from winreg import ConnectRegistry,OpenKey,OpenKeyEx, QueryValueEx,HKEY_CURRENT_USER
from ppt import OfficeApp

#================================================================================

class App:

    #-----------------------------------------------------------------------------------
    def ShowListFrame(self):
        self.listFrame.pack(fill=BOTH, side=TOP, padx=10, pady=0, expand=1)
        self.frameBottom.pack(fill=X, side=TOP, padx=10, pady=5, expand=0)

    #-----------------------------------------------------------------------------------
    def HideListFrame(self):
        self.listFrame.pack_forget()
        self.frameBottom.pack_forget()

    #-----------------------------------------------------------------------------------
    def PopulateListWithLinks(self):
        self.links = self.PowerPointApp.GetLinksFromSlides(
            int(self.root.winfo_x()) + int(self.root.winfo_width()/2), 
            int(self.root.winfo_y()) + int(self.root.winfo_height()/2)
        )

        #print('back to index')
        if len(self.links) != 0:
            for item in self.links:
                #print(dict(item))
                self.listbox.insert(END,str(item['link']))

            #print('filled list')
            self.frameTopLabelSrc.config(
                text= self.PowerPointApp.FileName,
            )
            #print('changed name')
            self.frameTopButtonFileOpen.config(
                text= 'Close file', 
                command= self.ClosePPT
            )

            self.ShowListFrame()
        #print('end populate')

    #-----------------------------------------------------------------------------------
    def ReplaceSelected(self):
        for item in self.listbox.curselection():
            text = self.listbox.get(item)
            text=text.replace(
                self.frameBottomFromInput.get(),
                self.frameBottomToInput.get()
            )
            #print(self.listbox.get(item), text)
            self.listbox.delete(item)
            self.listbox.insert(item, text)
            self.links[item]['link'] = text

        self.listbox.update()
        self.PowerPointApp.SetLinksInSlides(
            self.links, 
            int(self.root.winfo_x()) + int(self.root.winfo_width()/2), 
            int(self.root.winfo_y()) + int(self.root.winfo_height()/2)
        )

    #-----------------------------------------------------------------------------------
    def ClosePPT(self):
        if self.PowerPointApp.FileName != "":
            self.PowerPointApp.Close()
        self.listbox.delete(0,END)
        #print("Deleted list")

        self.frameTopLabelSrc.config(
            text= 'No file have been choosen'
        )

        self.frameTopButtonFileOpen.config(
            text= 'Choose File',
            command= self.PopulateListWithLinks
        )

        self.HideListFrame()
        #print("Reset File Name")

    #-----------------------------------------------------------------------------------
    def ExitApp(self):
        if messagebox.askokcancel(
            title = "Quit", 
            message = "This action will close the powerpoint file and exit the program. Do you want to quit now?", 
            parent= self.root):
            
            if self.PowerPointApp.FileName != "":
                self.PowerPointApp.Close()
                #print("Closed powerpoint")

            del self.frameBottomButton
            del self.frameBottomFromInput
            del self.frameBottomFromLabel
            del self.frameBottomToInput
            del self.frameBottomToLabel
            del self.frameBottom

            del self.frameTopLabelSrc
            del self.frameTopButtonFileOpen
            del self.frameTop

            del self.scrollbar
            del self.listbox
            del self.listFrame

            self.root.destroy()
            del self.root
            #print('Exiting now')
            return True
        #print('Didn\'t exit')

    #-----------------------------------------------------------------------------------
    def monitor_changes(self):
        registry = ConnectRegistry(None, HKEY_CURRENT_USER)
        key = OpenKey(registry, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize')
        mode = QueryValueEx(key, "AppsUseLightTheme")

        cnf = {
            'mainWindow': {
                'light': {'bg':'#ddd'}, 
                'dark': {'bg': '#0A2342'}
            },
            'frame': {
                'light': {'bg':'#ddd', 'fg':'#006ADF'}, 
                'dark': {'bg': '#0A2342', 'fg': '#E8E6E1'}
            },
            'scrollbar': {
                'light': {'width': 12, 'elementborderwidth': 5}, 
                'dark': {'width': 12, 'elementborderwidth': 5}
            },
            'label': {
                'light': {'bg':'#ddd', 'fg':'#000'}, 
                'dark': {'bg': '#0A2342', 'fg': '#fff'}
            },
            'input': {
                'light': {'bg':'#EEEEEE', 'fg':'#212121'}, 
                'dark': {'bg': '#798797', 'fg': '#FAFAFA'}
            },
            'button': {
                'light': {'bg':'#006ADF', 'fg':'#fff'}, 
                'dark': {'bg': '#006ADF', 'fg': '#fff'}
            },
            'list': {
                'light': {'bg': '#ccc','fg': '#000', 'selectbackground': '#FF0056', 'selectforeground':'#FFFFFF'},
                'dark': {'bg': '#07172A','fg': '#E8E6E1', 'selectbackground': '#A30037', 'selectforeground':'#FFFFFF'}
            }
        }

        if (self.root['bg']!='#ddd' and mode[0]) or (self.root['bg']!='#0A2342' and mode[0]==False):
            theme = 'light' if mode[0] else 'dark'

            color = cnf['mainWindow']['light'] if mode[0] else cnf['mainWindow']['dark']
            self.root.config(color)
            self.mainWindow.config(color)
            self.frameBottomGet.config(color)

            color = cnf['frame'][theme]
            self.frameBottom.config(color)
            self.frameTop.config(color)

            color = cnf['list'][theme]
            self.listbox.config(color)

            color = cnf['scrollbar'][theme]
            self.scrollbar.config(color)

            color = cnf['label'][theme]
            self.frameBottomToLabel.config(color)
            self.frameBottomFromLabel.config(color)
            self.frameTopLabelSrc.config(color)

            color = cnf['input'][theme]
            self.frameBottomFromInput.config(color)
            self.frameBottomToInput.config(color)

            color = cnf['button'][theme]
            self.frameTopButtonFileOpen.config(color)
            self.frameBottomButton.config(color)

        if len(self.listbox.curselection()) >= 1 and len(self.frameBottomFromInput.get())>0:
            self.frameBottomButton.config(state='normal')
        else:
            self.frameBottomButton.config(state=DISABLED)
        
        self.root.after(1000,self.monitor_changes)

    #-----------------------------------------------------------------------------------
    def __init__(self):
        # Initialize app container
        self.cnf = {
            'font': {'font': ('Inter','9', 'bold')},
        }

        self.root = Tk()
        self.root.title("Powerpoint Presentations Link Updater")
        self.root.geometry('650x350')
        self.root.minsize(650,325)
        self.root.iconbitmap(r'.\trendence_black.ico')

        self.PowerPointApp = OfficeApp("PowerPoint.Application")
        self.FileName = 'No file have been choosen yet'

        # Create Main Window
        self.mainWindow = Frame(self.root)
        self.mainWindow.pack_propagate(0)
        self.mainWindow.pack(fill=BOTH, expand=1, side=LEFT, pady=5)
        self.mainWindow.grid_rowconfigure(0, weight=1) 
        self.mainWindow.grid_columnconfigure(0, weight=1) 

        # Links listbox
        self.listFrame = Frame(self.mainWindow)

        self.listbox = Listbox(self.listFrame)
        self.listbox.pack(expand=1, fill=BOTH, side=LEFT)

        self.scrollbar = Scrollbar(self.listFrame)
        self.scrollbar.pack(fill=Y, side=LEFT)

        self.scrollbar.config(
            command = self.listbox.yview
        )

        self.listbox.config(
            selectmode=EXTENDED,
            exportselection= False,
            highlightthickness=0,
            yscrollcommand = self.scrollbar.set
        )
        self.listbox.config(self.cnf['font'])

        # File Frame
        self.frameTop = LabelFrame(self.mainWindow, text='Choose File')

        self.frameTopButtonFileOpen = Button(self.frameTop)
        self.frameTopButtonFileOpen.grid(in_=self.frameTop, column=0, row=0, padx=10, pady=5)

        self.frameTopLabelSrc = Label(self.frameTop)
        self.frameTopLabelSrc.grid(column=1, row=0, padx=10, pady=5)

        self.frameTopLabelSrc.config({'text': self.FileName})
        self.frameTopLabelSrc.config(self.cnf['font'])

        # Find and Replace Frame
        self.frameBottom = LabelFrame(self.mainWindow, text='Selected Links')

        self.frameBottomGet = Frame(self.frameBottom)
        self.frameBottomGet.pack(fill=X, side=TOP, padx=10, pady=0, expand=1)

        self.frameBottomFromLabel = Label(self.frameBottomGet, text='Replace')
        self.frameBottomFromLabel.pack(side=LEFT, padx=0, pady=0)
        self.frameBottomFromLabel.config(self.cnf['font'])

        self.frameBottomFromInput = Entry(self.frameBottomGet)
        self.frameBottomFromInput.pack(side=LEFT, fill=X, padx=0, pady=0, expand=1)
        self.frameBottomFromInput.config(self.cnf['font'])

        self.frameBottomToLabel = Label(self.frameBottomGet, text='With ')
        self.frameBottomToLabel.pack(side=LEFT, padx=0, pady=0)
        self.frameBottomToLabel.config(self.cnf['font'])

        self.frameBottomToInput = Entry(self.frameBottomGet)
        self.frameBottomToInput.pack(side=LEFT, fill=X, padx=0, pady=0, expand=1)
        self.frameBottomToInput.config(self.cnf['font'])

        self.frameBottomButton = Button(self.frameBottom)
        self.frameBottomButton.pack(fill=X, side=TOP, padx=10, pady=5, expand=1)
        self.frameBottomButton.config(text='Replace Selected', command= self.ReplaceSelected, height=1)
        self.frameBottomButton.config(self.cnf['font'])

        self.frameTop.pack(fill=X, side=TOP, padx=10, pady=5)
        self.frameTop.config(height=50)
        self.frameTop.config(self.cnf['font'])

        self.frameBottom.config(self.cnf['font'])

        self.frameTopButtonFileOpen.config(text= 'Choose a PowerPoint file', command= self.PopulateListWithLinks)
        self.frameTopButtonFileOpen.config(self.cnf['font'])

        self.root.protocol("WM_DELETE_WINDOW", self.ExitApp)
        self.monitor_changes()
        self.root.mainloop()

if __name__ == '__main__':
    LinkUpdater = App()
    r = 0
    LF=[]
