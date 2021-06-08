from Ui import Ui
from tkinter import Label, Checkbutton, BooleanVar, Button, Listbox, Entry, END, filedialog
from tkinter import ttk


def _state(state):
    return 'normal' if state else 'disabled'


class Structure(Ui):
    ALL = "Tout"
    CHOOSE_STRUCT = "Choississez la/les structure(s)"

    def __init__(self, frame, list_rec_box=None):
        """
        :param list_rec_box: liste des structures
        :type list_rec_box: list
        """
        super().__init__(frame, "Structure")
        if list_rec_box is None:
            list_rec_box = []
        self.list_rec_box = list_rec_box
        list_rec_box.insert(0, self.ALL)

        label_struct = Label(frame, text='Structure :', font=60)
        label_struct.grid(row=0, column=0, ipady=30, padx=5)

        chk_struct_val = BooleanVar(value=False)
        chk_struct = Checkbutton(frame, var=chk_struct_val, command=lambda: self.is_chk_struct())
        chk_struct.grid(row=1, column=1, padx=4)
        self.chk_struct_val = chk_struct_val
        self.chk_struct = chk_struct

        # combobox qui récupère les recycleries comme valeur
        combo_struct = ttk.Combobox(frame, values=list_rec_box, width=28, state='disabled')
        combo_struct.set(self.CHOOSE_STRUCT)  # valeur par défaut du combobox
        combo_struct.grid(row=1, column=2)
        self.combo_struct = combo_struct

        # bouton qui ajoute la recylerie sélectionnée du combobox dans la listbox
        btn_add = Button(frame, text='Ajouter', command=lambda: self.add_struct(), state='disabled')
        btn_add.grid(row=1, column=3)
        self.btn_add = btn_add

        # bouton qui supprime la recylerie sélectionnée dans la listbox
        btn_del = Button(frame, text='Supprimer', command=lambda: self.del_struct(), state='disabled')
        btn_del.grid(row=1, column=4, padx=5)
        self.btn_del = btn_del

        list_struct = Listbox(frame, width=50)  # listbox qui va contenir les recyleries à étudier
        list_struct.grid(row=1, column=5)
        self.list_struct = list_struct

        label_secteur = Label(frame, text='Secteur :', font=60)
        label_secteur.grid(row=2, column=0, ipady=70, padx=5)

        label_insee = Label(frame, text='Insee :')
        label_insee.grid(row=3, column=1)

        chk_insee_one = BooleanVar(value=False)
        chk_insee_many = BooleanVar(value=False)

        # si on coche, active la ligne pour l'insertion d'1 code sinon reste désactivé
        chk_one = Checkbutton(frame, text='1 code :', var=chk_insee_one, command=lambda: self.is_chk_one())
        chk_one.grid(row=3, column=2, padx=4)
        self.chk_insee_one = chk_insee_one
        self.chk_one = chk_one

        entry_insee = Entry(frame, state='disabled')  # désactivé tant que la checkbox n'est pas cochée
        entry_insee.grid(row=3, column=3, padx=10)
        self.entry_insee = entry_insee

        # si on coche, active la ligne pour l'insertion de plusieurs codes sinon reste désactivé
        chk_many = Checkbutton(frame, text='plusieurs codes :', var=chk_insee_many,
                               command=lambda: self.is_chk_many())
        chk_many.grid(row=4, column=2, padx=4)
        self.chk_insee_many = chk_insee_many
        self.chk_many = chk_many

        btn_insee = Button(frame, text='Importer votre fichier', command=lambda: self.file_insee(),
                           state='disabled')  # désactivé tant que la checkbox n'est pas cochée
        btn_insee.grid(row=4, column=3)
        self.btn_insee = btn_insee

        label_file = Label(frame, text='')
        label_file.grid(row=4, column=4)
        self.label_file = label_file

        list_insee = Listbox(frame, height=10)
        list_insee.grid(row=4, column=5, padx=40)
        self.list_insee = list_insee

        # passe à la prochaine fenêtre et prends en compte les données inscrites de la première fenêtre
        btn_next = Button(frame, text='Suivant', command=lambda: new_window())
        btn_next.grid(row=6, column=5, padx=40, pady=20)

    def is_chk_struct(self):
        """
        Change state of the structures' widgets
        """
        # state = 'normal' if self.chk_struct_val.get() else 'disabled'
        for s in [self.combo_struct, self.btn_add, self.btn_del]:
            s.config(state=_state(self.chk_struct_val.get()))

    def is_chk_one(self):
        """
        fonction qui change l'état des widgets en lien avec 1 code insee
        """
        self.entry_insee.config(state=_state(self.chk_insee_one.get()))
        self.chk_many.config(state=_state(not self.chk_insee_one.get()))

    def is_chk_many(self):
        """
        fonction qui change l'état des widgets en lien avec plusieurs codes insee
        """
        self.btn_insee.config(state=_state(self.chk_insee_many.get()))
        self.chk_one.config(state=_state(not self.chk_insee_many.get()))
        if not self.chk_insee_many.get():
            self.list_insee.delete(0, END)
            self.label_file.config(text='')

    def add_struct(self):
        """
        fonction pour le bouton ajouter des structures
        """
        list_files = self.list_struct.get(0, END)
        value = self.combo_struct.get()
        if self.combo_struct.get() == self.ALL:
            for struct in self.list_rec_box:
                if struct not in list_files:
                    self.list_struct.insert(END, struct)
            idx = self.list_struct.get(0, END).index(self.ALL)
            self.list_struct.delete(idx)

        if value not in list_files and value != self.ALL and value != self.CHOOSE_STRUCT:
            self.list_struct.insert(END, value)

    def del_struct(self):
        """
        fonction pour le bouton supprimer des structures
        """
        item_selected = self.list_struct.curselection()
        self.list_struct.delete(item_selected[0])

    def file_insee(self):
        """
        fonction pour chercher le fichier .txt contenant les plusieurs codes insee et l'insère dans la listbox
        """
        title = "Sélectionnez votre fichier"
        ft = (("Fichier Texte", "*.txt"), ("Fichier pdf", "*.pdf"))
        filename = filedialog.askopenfilename(initialdir="/", title=title, filetypes=ft)

        self.label_file.config(text=filename)
        fichier = open(str(filename), "r")
        lines = fichier.readlines()
        fichier.close()
        self.list_insee.delete(0, END)
        i = 1
        for line in lines:
            self.list_insee.insert(i, line.strip())
            i += 1
