from tkinter import *
import guru

root = Tk()
root.geometry('1920x1080')
root.configure(bg='#041f1e')
root.title("GRADING SNIPPET AND INVESTIGATOR - GURUNANAK COLLEGE OF ARTS AND SCIENCE")
def command():
   x=regno.get(1.0,END).split("\n")
   s=sheet.get()
   f=file.get()
   p=path.get()
   while("" in x):
      x.remove('')
   guru.funcGnc(x,s,f,p)
   for a,b in zip(guru.an,guru.ab):
         na=a
         re=b
         op_text.insert(INSERT, na+" "+re+"\n")
button = Button(root, text="OK", command=command)
head = Label(text = "GRADING SNIPPET AND INVESTIGATOR - GURUNANAK COLLEGE OF ARTS AND SCIENCE",font=20,fg="#000").place(x = 360,y = 20)
auth= Label(text = "AUTHOR : DINEESH R").place(x = 1190,y = 20)
regno = Text(root,width=50,height=20)
regno.place(x=170, y=300)
button.place(x=670, y=650)
sheet_name = Label(text = "SHEET NAME :").place(x = 40,y = 140) 
sheet = Entry(root)
sheet.place(x=170, y=140)
file_name = Label(text = "FILE NAME :").place(x = 40,y = 170)
file = Entry(root)
file.place(x=170, y=170)
path_ = Label(text = "PATH TO LOCATION :").place(x = 40,y = 200)
path = Entry(root)
path.place(x=170, y=200)
reg= Label(text = "REGISTRATION NUMBERS :").place(x = 40,y = 250)
opt= Label(text = "REFERENCE BOX").place(x = 950,y = 250)
op_text = Text(root, height=20, width=70)
op_text.place(x=720,y=300)
root.mainloop()