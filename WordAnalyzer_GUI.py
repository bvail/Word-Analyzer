#!/usr/bin/env python3 -u

# Note: the -u is runnig in unbuffered so the output shows up as the programs runs

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import messagebox

import WordAnalyzer

class Word_GUI:
        
        
        def __init__(self, master):
                
                # Window title
                master.title('Word Analyzer')
                
                # Does not allow resizing the window
                # master.resizable(False, False)

                # Creates own style, for all widgets
                self.style = ttk.Style()
                self.style.configure('TFrame', background = '#ededed')
                self.style.configure('TButton', background = '#ededed')
                self.style.configure('TLabel', background = '#ededed')
                
                # names self.master for func: Close_App
                self.master = master
                
        
                
                # Main top frame
                self.frame_header = ttk.Frame(master)
                self.frame_header.pack(fill = BOTH)
                
                self.logo = PhotoImage(file = 'WAlogo.png').subsample(2,2)
                
                # Labels in top frame
                ttk.Label(self.frame_header, image = self.logo).grid(row = 0, column = 0, sticky = 'w')
                ttk.Label(self.frame_header, text = 'Word Analyzer', font = ('Courier', 30)).grid(row = 0, column = 1, sticky = 'e')
                
                # Adds blank space in header
                self.frame_header.grid_columnconfigure(2, minsize = 230)
                self.frame_header.grid_rowconfigure(2, minsize = 5)
                
                
                
                
                # Separator Frame = Gave own frame because grid geometry used in header frame
                # Need pack geometry to fill
                self.separator_frame = ttk.Frame(master)
                self.separator_frame.pack(fill = BOTH)
                ttk.Separator(self.separator_frame, orient = 'horizontal').pack(fill = BOTH, pady = 5)
              
                
        
                
                # Main content frame
                self.frame_content = ttk.Frame(master)
                self.frame_content.pack(fill = BOTH, expand = 1)
                
                # Labels in content frame
                ttk.Label(self.frame_content, text = '* select .docx file only', foreground = 'gray').grid(row = 1, column = 1, columnspan = 2, sticky = 'nw')
                ttk.Label(self.frame_content, text = 'Total Word Count:').grid(row = 6, column = 0, padx = 20, sticky = 's')
                ttk.Label(self.frame_content, text = 'Unique Word Count:').grid(row = 8, column = 0, padx = 20, sticky = 's')
                self.total_word_count = ttk.Label(self.frame_content, text = 'xxxx') # total word count displayed here...
                self.total_word_count.grid(row = 7, column = 0, sticky = 'n')
                self.unique_word_count = ttk.Label(self.frame_content, text = 'xxxx') # unique word count displayed here...
                self.unique_word_count.grid(row = 9, column = 0, sticky = 'n')
                ttk.Label(self.frame_content, text = 'Frequency Analysis:')
                ttk.Label(self.frame_content, text = 'Search Single Word').grid(row = 19, column = 1, sticky = 'w')
                ttk.Label(self.frame_content, text = 'Count:').grid(row = 19, column = 2)
                self.single_word_count = ttk.Label(self.frame_content, text = 'xxxx') # single word count from search function
                self.single_word_count.grid(row = 20, column = 2)
                
                
                # File select bar in content frame
                self.file_entry = ttk.Entry(self.frame_content, width = 30)
                self.file_entry.grid(row = 0, column = 1, columnspan = 2, sticky = 'w')
                self.file_name = 'File name here...' # initializes file_name for Analyze_File button function
                self.file_entry.insert(0, self.file_name)
                self.file_entry.config(foreground = 'gray')
                
                
                # Buttons in content frame
                ttk.Button(self.frame_content, text = 'Open', command = self.File_Opener).grid(row = 0, column = 0)
                ttk.Button(self.frame_content, text = 'Analyze', command = self.Analyze_File).grid(row = 0, column = 3,padx = 10)
                ttk.Button(self.frame_content, text = 'Clear All', command = self.Clear_App).grid(row = 21, column = 1, pady = 10, sticky = 'w')
                ttk.Button(self.frame_content, text = 'Close', command = self.Close_App).grid(row = 21, column = 3, pady = 10, sticky = 'e')
                self.search_button = ttk.Button(self.frame_content, text = 'Search', state = DISABLED, command = self.Search_Word_List)
                self.search_button.grid(row = 20, column = 3, sticky = 'e')
                

                # Text box for results
                self.results = scrolledtext.ScrolledText(self.frame_content, width = 50, height = 30)
                self.results.grid(row = 3, column = 1, rowspan = 15, columnspan = 3, pady = 10, sticky = 'w')
                
                # Text box to input word search
                self.word_entry = ttk.Entry(self.frame_content, width = 20)
                self.word_entry.grid(row = 20, column = 1, sticky = 'w')
                self.single_word = 'Single word here...' #initializes sing_word for word_entry input bar
                self.word_entry.insert(0, self.single_word)
                self.word_entry.config(foreground = 'gray')
               
               
        
        # Func that opens file select window
        def File_Opener(self):
                self.file_name = filedialog.askopenfilename(initialdir = '/',
                                            filetypes = (('docx file', '*.docx'), ('all files', '*.*')),
                                            title = 'Choose a File:')
                if self.file_name != '':
                        self.file_entry.delete(0, END)
                        self.file_entry.config(foreground = 'black')
                        self.file_entry.insert(0, self.file_name)
                        # if self.file_name == '': print('True')
               
        # Fucntion to analyze the document
        def Analyze_File(self):
                
                # Makes sure the file is a docx
                if self.file_name.endswith('.docx') == False:
                        messagebox.showwarning(title = 'Error!',
                                            message = '.docx files are the only supported files.\n\n'
                                            '.doc, .pdf, .txt, etc. are not yet supported.')
                else:
                        # Calls Word Analyzer to analyze document
                        self.doc = WordAnalyzer.WordAnalyzer(self.file_name)
                        self.doc.analyze_doc()
                        self.word_list = self.doc.get_word_list()
                        
                        # Prints word_list in text box
                        for i in range (len(self.word_list)):
                                #self.results.insert(END, (self.word_list[i][0] + '  -  ' + str(self.word_list[i][1]) + '\n'))
                                string = ('{:.<25s}{:.>6s}\n' .format(self.word_list[i][0], str(self.word_list[i][1])))
                                self.results.insert(END, string)
                                
                                
                        
                        # Prints word_count
                        self.total_word_count.config(text = self.doc.word_count())
                        
                        #Prints unique_word_count
                        self.unique_word_count.config(text = self.doc.unique_word_count())
                        
                        # Enables search button for single word searches
                        self.search_button.config(state = NORMAL)
         
        
        # Function to search word list for specific word
        def Search_Word_List(self):
                
                self.single_word = self.word_entry.get()

                if self.single_word == 'Search single word...':
                        return
                
                self.single_word.lower()
                
                count = self.doc.search_word_list(self.single_word)
                
                
                
                # for i in range(len(self.word_list)):
                #         if self.single_word == self.word_list[i][0]:
                #                 self.single_word_count.config(text = self.word_list[i][1])
                #                 return
                
                if count != None:
                        self.single_word_count.config(text = count)
                else:
                        self.single_word_count.config(text = 0)   
        
                                      
        def Clear_App(self):
                
                # resets word count labels
                self.total_word_count.config(text = 'xxxx')
                self.unique_word_count.config(text = 'xxxx')
                self.single_word_count.config(text = 'xxxx')
                
                # resets file name entry window
                self.file_entry.delete(0, END)
                self.file_name = 'File name here...'
                self.file_entry.insert(0, self.file_name)
                self.file_entry.config(foreground = 'gray')
                
                # resets search single word entry window
                self.word_entry.delete(0, END)
                self.single_word = 'Single word here...'
                self.word_entry.insert(0, self.single_word)
                self.word_entry.config(foreground = 'gray')
                
                
                # resets results text box
                self.results.delete('1.0', 'end')
                
                # resets variables in function: Analyze_File
                self.doc = None
                self.word_list = None
                
                
        
        def Close_App(self):
                close = messagebox.askyesno(title = '', message = 'Are you sure you want to exit Word Analyzer? Unsaved data will be lost.')
                if close == True:
                        self.master.destroy()
                        
                     
        


def main():
        
        root = Tk()
        
        # style = ttk.Style()
        # style.theme_use('classic')
        
        Analyzer = Word_GUI(root)
        root.mainloop()


if __name__ == '__main__': main()