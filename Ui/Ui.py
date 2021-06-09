from tkinter import END


class Ui:
    ALL = "Tout"

    def __init__(self, frame, title):
        self.frame = frame
        frame.title(title)
        frame.minsize(400, 300)

    def add(self, t_select, combo, t_all, choose_text):
        """

        :param t_select: listbox containing all selected items
        :param combo: combo containing all items
        :param t_all:
        :param choose_text:
        :return:
        """
        list_files = t_select.get(0, END)
        value = combo.get()

        if value == self.ALL:
            for t in t_all:
                if t not in list_files:
                    t_select.insert(END, t)
            idx = t_select.get(0, END).index(self.ALL)
            t_select.delete(idx)
        elif value not in list_files and value != choose_text:
            t_select.insert(END, value)

    def delete(self, t_select):
        selected_item = t_select.curselection()
        t_select.delete(selected_item[0])
