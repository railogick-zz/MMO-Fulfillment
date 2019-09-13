import tkinter as tk
import tkinter.ttk as ttk


class MMO_App(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.parent.title("MMO Fulfillment")
        self.parent.grid_rowconfigure(1, weight=1)
        self.parent.grid_columnconfigure(0, weight=1)
        self.col_names = [('Full Name', 150), ('Address', 150), ('City', 100), ('State', 40), ('Zip', 80),
                          ('Phone', 90), ('Email Address', 180)]

        # Break main window into 3 sections
        self.top_frame = tk.Frame(self.parent, bg='cyan', height=50)
        self.center_frame = tk.Frame(self.parent, bg='white')
        self.bottom_frame = tk.Frame(self.parent, bg='gray', height=50)

        self.top_frame.grid(row=0, sticky='nsew')
        self.center_frame.grid(row=1, sticky='nsw')
        self.center_frame.columnconfigure(0, weight=1)
        self.bottom_frame.grid(row=2, sticky='new')

        # Data Entry Frame
        self.data_entry = {}
        self.data_label = {}
        self.data_frame = tk.Frame(self.center_frame, padx=8, pady=3)
        self.data_frame.grid(row=0, column=0, sticky='ns')
        self.data_frame.columnconfigure(0, weight=1)
        self.dnc_chkbox = tk.Checkbutton(self.data_frame, text="Do Not Contact")
        self.dnc_chkbox.grid(row=0, column=0, columnspan=2, sticky='ew')
        self.create_data(self.data_frame)
        self.submit_frame = tk.Frame(self.center_frame, padx=8, pady=2)
        self.submit_frame.grid(row=1, column=0, sticky='nsew')
        self.submit_frame.columnconfigure(0, weight=1)
        self.submit_frame.rowconfigure(0, weight=1)
        self.test_button = tk.Button(self.submit_frame, text='Submit')
        self.test_button.grid(row=1, column=0, padx=8, pady=5)


        # Tree view Frame.
        self.tree_frame = tk.Frame(self.center_frame, width=10, padx=5, pady=12)
        self.tree_frame.grid(row=0, column=1, rowspan=2, sticky='nsew')
        self.tree_frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview()
        self.create_tree(self.tree_frame)

    def create_data(self, frame_loc):
        # Generate Entry Widgets
        counter = 1
        for name, width in self.col_names:

            temp_lbl = tk.Label(frame_loc, text=f'{name}:', anchor='e', padx=5, pady=3)
            temp_lbl.grid(row=counter, column=0)
            self.data_label[name] = temp_lbl

            temp_entry = tk.Entry(self.data_frame)
            temp_entry.grid(row=counter, column=1)
            self.data_entry[name] = temp_entry

            counter += 1

    def create_tree(self, frame_loc):

        self.tree = ttk.Treeview(frame_loc, columns=[x for x, _ in self.col_names], show='headings')
        for name, width in self.col_names:
            self.tree.column(name, minwidth=0, width=width, anchor='center')
            self.tree.heading(name, text=name)
        self.tree.rowconfigure(0, weight=1)
        self.tree.pack(fill='x')

        # test data - REMOVE LATER
        self.tree.insert('', 'end', values=('John Doe', '123 Any Street', 'Anytown', 'Ohio', '44838-3333', '123-456-7890', 'email@email.com'))


def main():
    root = tk.Tk()
    app = MMO_App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
