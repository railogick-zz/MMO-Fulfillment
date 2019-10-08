import tkinter as tk
import tkinter.ttk as ttk

import mmo_import_xml


class MMO_App(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        parent.geometry('850x375')
        self.parent = parent
        self.parent.title("MMO Fulfillment")
        self.parent.grid_rowconfigure(1, weight=1)
        self.parent.grid_columnconfigure(0, weight=1)
        self.ent_data = {}
        self.lbl_data = {}
        self.data_entry_headers = [('Full Name', 150, 'w'), ('Address', 150, 'w'), ('City', 100, 'w'),
                                   ('State', 40, 'center'), ('Zip', 80, 'center'), ('Phone', 90, 'center'),
                                   ('Email Address', 180, 'center'), ('DNC', 40, 'center')]

        # Break main window into 3 sections
        self.frame_main_top = tk.Frame(self.parent, bg='cyan', height=50)
        self.frame_main_top.grid(row=0, sticky='nsew')
        self.frame_main_center = tk.Frame(self.parent, bg='white')
        self.frame_main_center.grid(row=1, sticky='nsew')
        self.frame_main_center.grid_rowconfigure(0, weight=1)
        self.frame_main_bottom = tk.Frame(self.parent, bg='gray', height=50)
        self.frame_main_bottom.grid(row=2, sticky='nsew')

        # --- Top Frame Contents ---
        self.btn_import = tk.Button(self.frame_main_top, text='Import XML Files')
        self.btn_import.grid(row=0, column=0)
        self.btn_data_entry = tk.Button(self.frame_main_top, text='Data Entry')
        self.btn_data_entry.grid(row=0, column=1)
        self.btn_process = tk.Button(self.frame_main_top, text='Process Data')
        self.btn_process.grid(row=0, column=2)
        self.btn_output = tk.Button(self.frame_main_top, text='Create Files')
        self.btn_output.grid(row=0, column=3)
        self.btn_close = tk.Button(self.frame_main_top, text='Quit')
        self.btn_close.grid(row=0, column=4)

        # --- Center Frame Contents ---

        self.notebook_treeview = ttk.Notebook(self.frame_main_center)
        self.tab_xml = tk.Frame(self.notebook_treeview)
        self.tab_xml.pack()
        self.notebook_treeview.add(self.tab_xml, text='XML Import')
        self.notebook_treeview.pack(expand=True, fill='both')
        self.tree_xml = ttk.Treeview()
        data_tree = [name for name, width, align in self.data_entry_headers]
        self.create_tree(self.tab_xml, data_tree)

        # Data Entry Frame
        # self.frame_data = tk.Frame(self.frame_main_center, padx=8, pady=3)
        # self.frame_data.grid(row=0, column=0, sticky='ns')
        # self.frame_data.columnconfigure(0, weight=1)
        # self.create_data(self.frame_data)
        #
        # # Submit Frame
        # self.frame_submit = tk.Frame(self.frame_main_center, padx=8, pady=2)
        # self.frame_submit.grid(row=1, column=0, sticky='nsew')
        # self.frame_submit.columnconfigure(0, weight=1)
        # self.frame_submit.rowconfigure(0, weight=1)
        # self.btn_test = tk.Button(self.frame_submit, text='Submit')
        # self.btn_test.grid(row=1, column=0, padx=8, pady=5)

        # Tree view Frame.

        # --- Bottom Frame Contents ---

    def create_data(self, location):
        # Do Not Contact Box
        self.chk_dnc = tk.Checkbutton(self.frame_data, text="Do Not Contact")
        self.chk_dnc.grid(row=0, column=0, columnspan=2, sticky='ew')
        # Generate Entry Widgets
        counter = 1
        for name, width, align in self.data_entry_headers:

            lbl_temp = tk.Label(location, text=f'{name}:', anchor='e', padx=5, pady=3)
            lbl_temp.grid(row=counter, column=0)
            self.lbl_data[name] = lbl_temp

            ent_temp = tk.Entry(self.frame_data)
            ent_temp.grid(row=counter, column=1)
            self.ent_data[name] = ent_temp

            counter += 1

    def create_tree(self, location, tree_headers):
        canvas_f1 = tk.Canvas(location)
        canvas_f1.pack(expand=True, fill='both')
        canvas_f2 = tk.Canvas(location)
        canvas_f2.pack(side='bottom', expand=False, fill='x')
        self.tree_xml = ttk.Treeview(canvas_f1, columns=tree_headers,
                                     show='headings')

        for name, width, align in self.data_entry_headers:
            self.tree_xml.column(name, width=width, anchor=align)
            self.tree_xml.heading(name, text=name)
        self.tree_xml.pack(side='left', expand=False, fill='y')

        self.tree_xml_scrollbar = tk.Scrollbar(canvas_f2, orient='horizontal', command=self.tree_xml.xview)
        self.tree_xml_scrollbar.pack(side='bottom', expand=True, fill='x')

        canvas_f2.configure(xscrollcommand=self.tree_xml_scrollbar.set)

        # test data - REMOVE LATER
        self.tree_xml.insert('', 'end', values=('John Doe', '123 Any Street', 'Anytown', 'Ohio', '44838-3333',
                                            '123-456-7890', 'email@email.com'))


def main():
    root = tk.Tk()
    app = MMO_App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
