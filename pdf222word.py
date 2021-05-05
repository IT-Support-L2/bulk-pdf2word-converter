import threading
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import *
import os
from pdf2docx import Converter
import webbrowser


class App:
    def __init__(self, root):
        
        _bgcolor = '#d9d9d9'
        _fgcolor = '#000000'
        _compcolor = '#d9d9d9'
        _ana1color = '#d9d9d9'
        _ana2color = '#ececec'
        self.style = ttk.Style()
        

        root.geometry("700x600+510+179")
        root.resizable(0,  0)
        root.title("iTech®")
        root.configure(background="#d9d9d9")
        root.configure(highlightbackground="#d9d9d9")
        root.configure(highlightcolor="black")
        self.dir = StringVar()
        self.title = tk.Label(root)
        self.title.place(relx=0.186, rely=0.033, height=30, width=437)
        self.title.configure(activebackground="#f9f9f9")
        self.title.configure(activeforeground="#000000")
        self.title.configure(background="#d9d9d9")
        self.title.configure(disabledforeground="#a3a3a3")
        self.title.configure(font="-family {Arial Rounded MT Bold} -size 10")
        self.title.configure(foreground="#000000")
        self.title.configure(highlightbackground="#d9d9d9")
        self.title.configure(highlightcolor="black")
        self.title.configure(text='''Bulk PDF2WORD Converter''')

        self.style.configure('TNotebook.Tab', background=_bgcolor)
        self.style.configure('TNotebook.Tab', foreground=_fgcolor)
        self.style.map('TNotebook.Tab', background=
            [('selected', _compcolor), ('active',_ana2color)])
        self.TNotebook1 = ttk.Notebook(root)
        self.TNotebook1.place(relx=0.057, rely=0.15, relheight=0.818
                , relwidth=0.891)
        self.TNotebook1.configure(takefocus="")
        self.TNotebook1_t1 = tk.Frame(self.TNotebook1)
        self.TNotebook1.add(self.TNotebook1_t1, padding=3)
        self.TNotebook1.tab(0, text="Converter", compound="left", underline="-1"
                ,)
        self.TNotebook1_t1.configure(background="#d9d9d9")
        self.TNotebook1_t1.configure(highlightbackground="#d9d9d9")
        self.TNotebook1_t1.configure(highlightcolor="black")
        self.TNotebook1_t2 = tk.Frame(self.TNotebook1)
        self.TNotebook1.add(self.TNotebook1_t2, padding=3)
        self.TNotebook1.tab(1, text="Help",compound="left",underline="-1",)
        self.TNotebook1_t2.configure(background="#d9d9d9")
        self.TNotebook1_t2.configure(highlightbackground="#d9d9d9")
        self.TNotebook1_t2.configure(highlightcolor="black")
        self.TNotebook1_t3 = tk.Frame(self.TNotebook1)
        self.TNotebook1.add(self.TNotebook1_t3, padding=3)
        self.TNotebook1.tab(2, text="About",compound="none",underline="-1",)
        self.TNotebook1_t3.configure(background="#d9d9d9")
        self.TNotebook1_t3.configure(highlightbackground="#d9d9d9")
        self.TNotebook1_t3.configure(highlightcolor="black")

        self.Label2 = tk.Label(self.TNotebook1_t1)
        self.Label2.place(relx=0.242, rely=0.066, height=33, width=327)
        self.Label2.configure(activebackground="#f9f9f9")
        self.Label2.configure(activeforeground="black")
        self.Label2.configure(background="#d9d9d9")
        self.Label2.configure(disabledforeground="#a3a3a3")
        self.Label2.configure(font="-family {Arial Rounded MT Bold} -size 9")
        self.Label2.configure(foreground="#000000")
        self.Label2.configure(highlightbackground="#d9d9d9")
        self.Label2.configure(highlightcolor="black")
        self.Label2.configure(text='''Enter PDF Files Directory''')

        self.start_btn = tk.Button(self.TNotebook1_t1)
        self.start_btn.place(relx=0.161, rely=0.418, height=42, width=431)
        self.start_btn.configure(activebackground="#ececec")
        self.start_btn.configure(activeforeground="#000000")
        self.start_btn.configure(background="#d9d9d9")
        self.start_btn.configure(disabledforeground="#a3a3a3")
        self.start_btn.configure(font="-family {Arial Rounded MT Bold} -size 9")
        self.start_btn.configure(foreground="#000000")
        self.start_btn.configure(highlightbackground="#d9d9d9")
        self.start_btn.configure(highlightcolor="black")
        self.start_btn.configure(pady="0")
        self.start_btn.configure(state='active')
        self.start_btn.configure(command=self._resetbutton())

        self.output_msg = tk.Message(self.TNotebook1_t1)
        self.output_msg.place(relx=0.160, rely=0.593, relheight=0.345
                , relwidth=0.697)
        self.output_msg.configure(background="#d9d9d9")
        self.output_msg.configure(font="-family {Arial} -size 9")
        self.output_msg.configure(foreground="#000000")
        self.output_msg.configure(highlightbackground="#d9d9d9")
        self.output_msg.configure(highlightcolor="black")
        self.output_msg.configure(justify='center')
        self.output_msg.configure(width=434)

        self.Entry1 = tk.Entry(self.TNotebook1_t1)
        self.Entry1.place(relx=0.161, rely=0.264, height=36, relwidth=0.7)
        self.Entry1.configure(background="white")
        self.Entry1.configure(disabledforeground="#a3a3a3")
        self.Entry1.configure(font="-family {Arial} -size 10")
        self.Entry1.configure(foreground="#000000")
        self.Entry1.configure(insertbackground="black")
        self.Entry1.configure(justify='center')
        self.Entry1.configure(textvariable=self.dir)

        self.help_text = ScrolledText(self.TNotebook1_t2)
        self.help_text.place(relx=0.048, rely=0.066, relheight=0.862
                , relwidth=0.911)
        self.help_text.configure(background="white")
        self.help_text.configure(font="TkTextFont")
        self.help_text.configure(foreground="black")
        self.help_text.configure(highlightbackground="#d9d9d9")
        self.help_text.configure(highlightcolor="black")
        self.help_text.configure(insertbackground="black")
        self.help_text.configure(insertborderwidth="3")
        self.help_text.configure(selectbackground="blue")
        self.help_text.configure(selectforeground="white")
        self.help_text.configure(wrap="none")
        self.help_text.insert(tk.END, ''' 
           Bulk pdf2word let you convert mutliple pdf file(s) with high quality and 100% precision in few seconds.





           - What should I input in Files Directory Entry?

             You should input the directory of your pdf files folder.


          - Correct input: C:/Users/user/Desktop/folder/

          * Important: A front-slash must be at the end of the directory!

          - Wrong input: C:\\Users\\user\\Desktop\\folder\\


          - Best practice:


           Put your pdf files you are willing to convert in a single folder,

           then input its directory exactly like explained above.


           Each doc file will take the filename of its correspondent pdf file.
        ''')
        self.help_text.configure(state='disabled')

        self.about_text = tk.Label(self.TNotebook1_t3)
        self.about_text.place(relx=0.097, rely=0.132, height=117, width=488)
        self.about_text.configure(activebackground="#f9f9f9")
        self.about_text.configure(activeforeground="black")
        self.about_text.configure(background="#d9d9d9")
        self.about_text.configure(disabledforeground="#a3a3a3")
        self.about_text.configure(font="-family {Arial} -size 9")
        self.about_text.configure(foreground="#000000")
        self.about_text.configure(highlightbackground="#d9d9d9")
        self.about_text.configure(highlightcolor="black")
        self.about_text.configure(text=''' 
        
        Bulk Pdf2Word Converter coded with Python 3.9.0 and tkinter.

        GUI designed with PAGE but I modified most part of the generated code.
        
        By iTech® 2021
        ''')

        self.linkedin_btn = tk.Button(self.TNotebook1_t3)
        self.linkedin_btn.place(relx=0.242, rely=0.549, height=42, width=318)
        self.linkedin_btn.configure(activebackground="#ececec")
        self.linkedin_btn.configure(activeforeground="#000000")
        self.linkedin_btn.configure(background="#d9d9d9")
        self.linkedin_btn.configure(disabledforeground="#a3a3a3")
        self.linkedin_btn.configure(font="-family {Arial Rounded MT Bold} -size 9")
        self.linkedin_btn.configure(foreground="#000000")
        self.linkedin_btn.configure(highlightbackground="#d9d9d9")
        self.linkedin_btn.configure(highlightcolor="black")
        self.linkedin_btn.configure(pady="0")
        self.linkedin_btn.configure(text='''LinkedIn''')
        self.linkedin_btn.configure(command=lambda:self.linked_in())

        self.github_btn = tk.Button(self.TNotebook1_t3)
        self.github_btn.place(relx=0.242, rely=0.703, height=42, width=318)
        self.github_btn.configure(activebackground="#ececec")
        self.github_btn.configure(activeforeground="#000000")
        self.github_btn.configure(background="#d9d9d9")
        self.github_btn.configure(disabledforeground="#a3a3a3")
        self.github_btn.configure(font="-family {Arial Rounded MT Bold} -size 9")
        self.github_btn.configure(foreground="#000000")
        self.github_btn.configure(highlightbackground="#d9d9d9")
        self.github_btn.configure(highlightcolor="black")
        self.github_btn.configure(pady="0")
        self.github_btn.configure(text='''Github''')
        self.github_btn.configure(command=lambda:self.git_hub())

        self.website_btn = tk.Button(self.TNotebook1_t3)
        self.website_btn.place(relx=0.242, rely=0.857, height=42, width=318)
        self.website_btn.configure(activebackground="#ececec")
        self.website_btn.configure(activeforeground="#000000")
        self.website_btn.configure(background="#d9d9d9")
        self.website_btn.configure(disabledforeground="#a3a3a3")
        self.website_btn.configure(font="-family {Arial Rounded MT Bold} -size 9")
        self.website_btn.configure(foreground="#000000")
        self.website_btn.configure(highlightbackground="#d9d9d9")
        self.website_btn.configure(highlightcolor="black")
        self.website_btn.configure(pady="0")
        self.website_btn.configure(text='''Portfolio Website''')
        self.website_btn.configure(command=lambda:self.portfolio())

    def pdfTOword(self):

        try:
            path = self.dir.get()
            filenames = os.listdir(path)
            extension = ".pdf"

            for file in range(len(filenames)):
                if filenames[file].endswith(extension):
                    for f in filenames:            
                        pdf_files = path+f
                        docx_files = path+f+'.docx'
                        cv = Converter(pdf_files)
                        self.output_msg.configure(text='''Processing...Please wait...''')
                        cv.convert(docx_files, start=0, end=None)
                        cv.close()
                        self.output_msg.configure(text='''Done!''')
                        self.start_btn.config(text="Start", command=self.startthread)

        except OSError:
            self.output_msg.configure(text="Looks like you had wrong input Please check your inputs and try again!")
            
        except Exception:
            self.output_msg.configure(text="Unexpected error occured! Please try again")
            
    def linked_in(self):
        url='https://linkedin.com/in/cyber-services'
        webbrowser.open_new_tab(url)

    def git_hub(self):
        url='https://github.com/IT-Support-L2'
        webbrowser.open_new_tab(url)

    def portfolio(self):
        url='https://hamdi-bouaskar.herokuapp.com'
        webbrowser.open_new_tab(url)

    def _resetbutton(self):
        self.running = False
        self.start_btn.config(text="Start", command=self.startthread)

    def startthread(self):
        self.running = True
        newthread = threading.Thread(target=self.StartTask)
        newthread.start()
        self.start_btn.configure(text="Stop", command=self._resetbutton)

    def StartTask(self):
        if self.running:
            self.pdfTOword()         

class AutoScroll(object):
    def __init__(self, master):
        try:
            vsb = ttk.Scrollbar(master, orient='vertical', command=self.yview)
        except:
            pass
        hsb = ttk.Scrollbar(master, orient='horizontal', command=self.xview)
        try:
            self.configure(yscrollcommand=self._autoscroll(vsb))
        except:
            pass
        self.configure(xscrollcommand=self._autoscroll(hsb))
        self.grid(column=0, row=0, sticky='nsew')
        try:
            vsb.grid(column=1, row=0, sticky='ns')
        except:
            pass
        hsb.grid(column=0, row=1, sticky='ew')
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(0, weight=1)
        # Copy geometry methods of master  (taken from ScrolledText.py)
        
        methods = tk.Pack.__dict__.keys() | tk.Grid.__dict__.keys() | tk.Place.__dict__.keys()
        
        for meth in methods:
            if meth[0] != '_' and meth not in ('config', 'configure'):
                setattr(self, meth, getattr(master, meth))

    @staticmethod
    def _autoscroll(sbar):
        
        def wrapped(first, last):
            first, last = float(first), float(last)
            if first <= 0 and last >= 1:
                sbar.grid_remove()
            else:
                sbar.grid()
            sbar.set(first, last)
        return wrapped

    def __str__(self):
        return str(self.master)

def _create_container(func):
    
    def wrapped(cls, master, **kw):
        container = ttk.Frame(master)
        container.bind('<Enter>', lambda e: _bound_to_mousewheel(e, container))
        container.bind('<Leave>', lambda e: _unbound_to_mousewheel(e, container))
        return func(cls, container, **kw)
    return wrapped

class ScrolledText(AutoScroll, tk.Text):
    
    @_create_container
    def __init__(self, master, **kw):
        tk.Text.__init__(self, master, **kw)
        AutoScroll.__init__(self, master)

import platform
def _bound_to_mousewheel(event, widget):
    child = widget.winfo_children()[0]
    if platform.system() == 'Windows' or platform.system() == 'Darwin':
        child.bind_all('<MouseWheel>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Shift-MouseWheel>', lambda e: _on_shiftmouse(e, child))
    else:
        child.bind_all('<Button-4>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Button-5>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Shift-Button-4>', lambda e: _on_shiftmouse(e, child))
        child.bind_all('<Shift-Button-5>', lambda e: _on_shiftmouse(e, child))

def _unbound_to_mousewheel(event, widget):
    if platform.system() == 'Windows' or platform.system() == 'Darwin':
        widget.unbind_all('<MouseWheel>')
        widget.unbind_all('<Shift-MouseWheel>')
    else:
        widget.unbind_all('<Button-4>')
        widget.unbind_all('<Button-5>')
        widget.unbind_all('<Shift-Button-4>')
        widget.unbind_all('<Shift-Button-5>')

def _on_mousewheel(event, widget):
    if platform.system() == 'Windows':
        widget.yview_scroll(-1*int(event.delta/120),'units')
    elif platform.system() == 'Darwin':
        widget.yview_scroll(-1*int(event.delta),'units')
    else:
        if event.num == 4:
            widget.yview_scroll(-1, 'units')
        elif event.num == 5:
            widget.yview_scroll(1, 'units')

def _on_shiftmouse(event, widget):
    if platform.system() == 'Windows':
        widget.xview_scroll(-1*int(event.delta/120), 'units')
    elif platform.system() == 'Darwin':
        widget.xview_scroll(-1*int(event.delta), 'units')
    else:
        if event.num == 4:
            widget.xview_scroll(-1, 'units')
        elif event.num == 5:
            widget.xview_scroll(1, 'units')

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()





