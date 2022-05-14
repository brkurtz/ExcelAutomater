from EXCELgui import *
'''
Idea from chapter 13 from Automate the boring stuff https://automatetheboringstuff.com/2e/chapter13/ used it to build
framework and base file to start then added a gui and made it print results to a new excel sheet instead of printing it 
to a pyfile to use as quick access for the file.
'''

def main():

    """
    Window controls
    """
    window = Tk()
    window.title('Census Sheet Consolidation')
    window.geometry('600x200')
    window.resizable(False, False)
    widgets = Excelgui(window)
    window.mainloop()


if __name__ == '__main__':
    main()
