from tkinter import ttk, Tk
from tkinter import *
import time, threading


def update(pb, root, value_label):
    for x in range(10):
        
        pb['value']+=10
        root.update()
        time.sleep(0.5)
        value_label['text'] = f"Current Progress: {pb['value']}%"
    root.destroy()
    print('after destroy')
    

    
def main():
    root=Tk()
    root.geometry('300x120')
    root.title('Progressbar Demo')
    
    pb = ttk.Progressbar(root, orient='horizontal', mode='determinate', length=280)
    pb.grid(column=0, row=0, columnspan=2, padx=10, pady=20)
    
    value_label = ttk.Label(root, text= f"Current Progress: {pb['value']}%")
    value_label.grid(column=0, row=1, columnspan=2)
    pb['value']=10

    
    t1 = threading.Thread(target=update, args=(pb, root, value_label))
    t1.daemon = True
    t1.start()

    root.mainloop()
    print('after main')
  

if __name__ == "__main__":
    main()