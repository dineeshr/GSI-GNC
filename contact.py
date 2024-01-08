from tkinter import *
from PIL import ImageTk, Image
def cont():
    cont = Tk()
    cont.geometry('1920x1080')
    cont.configure(bg='#041f1e')
    cont.title("GRADING SNIPPET AND INVESTIGATOR - GURUNANAK COLLEGE OF ARTS AND SCIENCE")
    heading = Label(cont,text = "CONTACT",font=20,fg="#000").place(x = 640,y = 20)
    p1=Label(cont,text = """AUTHOR : DINEESH R \n\n E-MAIL : dineeshdina23@gmail.com \n\n GITHUB : https://github.com/dineeshr""",font=90,fg="#000").place(x = 550,y = 320)
    cont.mainloop()