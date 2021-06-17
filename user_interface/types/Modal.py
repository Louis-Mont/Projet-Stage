from tkinter import Checkbutton, BooleanVar


class Modal:
    """
    Generates a Modal that mimics a checkbutton and gives access to a BooleanVar and a CheckButton
    """

    def __init__(self, frame, text, val, command=None):
        """
        :type val: bool
        """
        self.chk_var = BooleanVar(value=val)
        self.chk_btn = Checkbutton(frame, text=text, var=self.chk_var, command=command)
