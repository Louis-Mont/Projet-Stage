import datetime
from Ui import Ui
from tkinter import Label, Button, Listbox
from tkinter import ttk

from user_interface.types.Modal import Modal


def dates(ll, ul):
    return [str(i).zfill(2) for i in range(ll, ul)]


class Extract(Ui):
    CHOOSE_CAT = "Choississez la/les catégorie(s)"
    CHOOSE_MODAL = "Choississez la/les modalité(s)"

    def date_combo(self, values, row, col, start_val):
        cbb = ttk.Combobox(self.frame, values=values)
        cbb.grid(row=row, column=col)
        cbb.set(start_val)
        return cbb

    def __init__(self, frame, y_bdd, export, list_struct=None, list_cat=None, list_modal_combo=None):
        """
        :param y_bdd: the lowest year present in the BDD
        :type y_bdd: int
        :param export: The function used to extract the data when the button "Extract" is clicked
        :type export: function
        :param list_struct: liste des structures
        :type list_struct: list
        :param list_cat: liste des categories
        :type list_cat: list
        :param list_modal_combo: liste des origines des arrivages, sans doublons
        :type list_modal_combo: list
        """
        super().__init__(frame, "Options d'importation")

        if list_cat is None:
            list_cat = []
        list_cat.insert(0, self.ALL)

        if list_modal_combo is None:
            list_modal_combo = []
        list_modal_combo.insert(0, self.ALL)

        # Partie catégorie :
        label_cat = Label(frame, text='Catégorie :', font=60)
        label_cat.grid(row=0, column=0, ipady=30)

        combo_cat = ttk.Combobox(frame, values=list_cat, width=29)
        combo_cat.set(self.CHOOSE_CAT)
        combo_cat.grid(row=0, column=1)
        combo_cat.bind("<<ComboboxSelected>>", lambda e: combo_cat.get())

        btn_add = Button(frame, text='Ajouter',
                         command=lambda: self.add(list_cat, combo_cat, list_cat, self.CHOOSE_CAT))
        btn_add.grid(row=0, column=2)

        btn_del = Button(frame, text='Supprimer', command=lambda: self.delete(list_cat))
        btn_del.grid(row=0, column=3, padx=5)

        list_cat = Listbox(frame, width=50)
        list_cat.grid(row=0, column=5)

        # Partie temps :
        label_time = Label(frame, text='Temps :', font=60)
        label_time.grid(row=1, column=0, ipady=10)

        days = dates(1, 32)
        months = dates(1, 13)
        year = int(datetime.date.today().year)
        years = [i for i in range(y_bdd, year + 1)]

        label_start = Label(frame, text="Choisissez la date de début :")
        label_start.grid(row=2, column=1, padx=10)
        label_end = Label(frame, text="Choisissez la date de fin :")
        label_end.grid(row=3, column=1, pady=10)
        choose_day_start = self.date_combo(days, 2, 2, '01')
        choose_month_start = self.date_combo(months, 2, 3, '01')
        choose_year_start = self.date_combo(years, 2, 4, y_bdd)
        choose_day_end = self.date_combo(days, 3, 2, '31')
        choose_month_end = self.date_combo(months, 3, 3, 12)
        choose_year_end = self.date_combo(years, 3, 4, year)

        # Partie quantitative :
        label_qty = Label(frame, text='Quantitative :', font=60)
        label_qty.grid(row=4, column=0, ipady=40, padx=5)

        label_modal = Label(frame, text='Modalité(s) :')
        self.label_modal = label_modal

        combo_modal = ttk.Combobox(frame, values=list_modal_combo, width=25)
        combo_modal.set(self.CHOOSE_MODAL)
        self.combo_modal = combo_modal

        btn_add_modal = Button(frame, text='Ajouter',
                               command=lambda: self.add(listbox_modal, combo_modal, list_modal_combo,
                                                        self.CHOOSE_MODAL))
        self.btn_add_modal = btn_add_modal

        btn_del_modal = Button(frame, text='Supprimer', command=lambda: self.delete(listbox_modal))
        self.btn_del_modal = btn_del_modal

        listbox_modal = Listbox(frame, width=30)
        self.listbox_modal = listbox_modal

        # elem : col
        self._opts = {
            label_modal: 1,
            combo_modal: 2,
            btn_add_modal: 3,
            btn_del_modal: 4,
            listbox_modal: 5
        }

        w_coll = Modal(frame, 'Poids collecté (en Kg)', False)
        w_sold = Modal(frame, 'Poids vendu (en Kg)', False)
        revenue = Modal(frame, 'Chiffre d\'affaire (en €)', False)
        n_coll = Modal(frame, 'Nombre d\'objet collecté', False)
        n_sold = Modal(frame, 'Nombre d\'objet vendu', False)

        w_coll.chk_btn.configure(command=lambda: self.show_modal_opt(w_coll.chk_var.get()))
        w_coll.chk_btn.grid(row=5, column=1)

        w_sold.chk_btn.grid(row=5, column=2)

        revenue.chk_btn.grid(row=5, column=3)

        n_coll.chk_btn.configure(command=lambda: self.show_modal_opt(n_coll.chk_var.get()))
        n_coll.chk_btn.grid(row=6, column=1)

        n_sold.chk_btn.grid(row=6, column=2)

        btn_export = Button(frame, text='Exporter',
                            command=lambda: export(frame, list_cat, choose_day_start, choose_month_start,
                                                   choose_year_start, choose_day_end, choose_month_end, choose_year_end,
                                                   list_struct, w_coll.chk_var, w_sold.chk_var, revenue.chk_var,
                                                   n_coll.chk_var, n_sold.chk_var, listbox_modal))
        btn_export.grid(row=10, column=5, padx=40, pady=20)

    def show_modal_opt(self, show):
        if show:
            for k, v in self._opts.items():
                k.grid(row=7, column=v)
        else:
            for k, v in self._opts.items():
                k.grid_forget()
