from tkinter import Tk

from Ui.Structure import Structure


def next_window():
    pass


if __name__ == '__main__':
    struct_frame = Tk()
    extract_frame = Tk()
    struct = Structure(struct_frame, lambda: next_window(), )
