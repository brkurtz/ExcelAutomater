from EXCELgui import *

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
