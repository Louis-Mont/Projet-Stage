import datetime
from Ui import Ui
from tkinter import Label, Button, Listbox, END, Checkbutton, BooleanVar
from tkinter import ttk


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

    def __init__(self, frame, y_bdd, export, list_struct=None, list_cat_box=None, list_modal_combo=None):
        super().__init__(frame, "Options d'importation")

        if list_cat_box is None:
            list_cat_box = []
        list_cat_box.insert(0, self.ALL)

        if list_modal_combo is None:
            list_modal_combo = []
        list_modal_combo.insert(0, self.ALL)

        # Partie catégorie :
        label_cat = Label(frame, text='Catégorie :', font=60)
        label_cat.grid(row=0, column=0, ipady=30)

        combo_cat = ttk.Combobox(frame, values=list_cat_box, width=29)
        combo_cat.set(self.CHOOSE_CAT)
        combo_cat.grid(row=0, column=1)
        combo_cat.bind("<<ComboboxSelected>>", lambda e: combo_cat.get())

        btn_add = Button(frame, text='Ajouter',
                         command=lambda: self.add(list_cat, combo_cat, list_cat_box, self.CHOOSE_CAT))
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
        years = [i for i in range(int(y_bdd), year + 1)]

        label_start = Label(frame, text="Choisissez la date de début :")
        label_start.grid(row=2, column=1, padx=10)
        label_end = Label(frame, text="Choisissez la date de fin :")
        label_end.grid(row=3, column=1, pady=10)
        choose_day_start = self.date_combo(days, 2, 2, '01')
        choose_month_start = self.date_combo(months, 2, 3, '01')
        choose_year_start = self.date_combo(years, 2, 4, int(y_bdd))
        choose_day_end = self.date_combo(days, 3, 2, '31')
        choose_month_end = self.date_combo(months, 3, 3, 12)
        choose_year_end = self.date_combo(years, 3, 4, year)

        # Partie quantitative :
        label_qty = Label(frame, text='Quantitative :', font=60)
        label_qty.grid(row=4, column=0, ipady=40, padx=5)

        label_modal = Label(frame, text='Modalité(s) :')
        combo_modal = ttk.Combobox(frame, values=list_modal_combo, width=25)
        combo_modal.set(self.CHOOSE_MODAL)
        btn_add_modal = Button(frame, text='Ajouter',
                               command=lambda: self.add(list_box_modal, combo_modal, list_modal_combo,
                                                        self.CHOOSE_MODAL))
        btn_del_modal = Button(frame, text='Supprimer', command=lambda: self.delete(list_box_modal))
        list_box_modal = Listbox(frame, width=30)

        chk_val_w_coll = BooleanVar(value=False)
        chk_val_w_sold = BooleanVar(value=False)
        chk_val_revenue = BooleanVar(value=False)
        chk_val_n_coll = BooleanVar(value=False)
        chk_val_n_sold = BooleanVar(value=False)

        chkbox1 = Checkbutton(frame, text='Poids collecté (en Kg)', var=chk_val_w_coll,
                              command=lambda: self.modalite_collect(chk_val_w_coll, chk_val_n_coll, combo_modal,
                                                                    label_modal,
                                                                    btn_add_modal, btn_del_modal, list_box_modal))
        chkbox1.grid(row=5, column=1)
        chkbox2 = Checkbutton(frame, text='Poids vendu (en Kg)', var=chk_val_w_sold)
        chkbox2.grid(row=5, column=2)
        chkbox3 = Checkbutton(frame, text='Chiffre d\'affaire (en €)', var=chk_val_revenue)
        chkbox3.grid(row=5, column=3)
        chkbox4 = Checkbutton(frame, text='Nombre d\'objet collecté', var=chk_val_n_coll,
                              command=lambda: self.modalite_collect(chk_val_w_coll, chk_val_n_coll, combo_modal,
                                                                    label_modal,
                                                                    btn_add_modal, btn_del_modal, list_box_modal))
        chkbox4.grid(row=6, column=1)
        chkbox5 = Checkbutton(frame, text='Nombre d\'objet vendu', var=chk_val_n_sold)
        chkbox5.grid(row=6, column=2)

        btn_export = Button(frame, text='Exporter',
                            command=lambda: export(frame, list_cat, choose_day_start, choose_month_start,
                                                   choose_year_start,
                                                   choose_day_end,
                                                   choose_month_end, choose_year_end, list_struct, chk_val_w_coll,
                                                   chk_val_w_sold,
                                                   chk_val_revenue,
                                                   chk_val_n_coll, chk_val_n_sold, list_box_modal))
        btn_export.grid(row=10, column=5, padx=40, pady=20)

    def modalite_collect(self, chk_value1, chk_value4, combo_modalite, label_modalite, btn_add_modal, btn_del_modal,
                         list_box_modal):
        """
        fonction qui affiche les widgets en lien avec les modalités
        si les checkbox des collectes sont cochées sinon les cachent
        """
        if chk_value1.get():
            combo_modalite.grid(row=7, column=2)
            label_modalite.grid(row=7, column=1)
            btn_add_modal.grid(row=7, column=3)
            btn_del_modal.grid(row=7, column=4)
            list_box_modal.grid(row=7, column=5)
        elif chk_value4.get():
            combo_modalite.grid(row=7, column=2)
            label_modalite.grid(row=7, column=1)
            btn_add_modal.grid(row=7, column=3)
            btn_del_modal.grid(row=7, column=4)
            list_box_modal.grid(row=7, column=5)
        else:
            combo_modalite.grid_forget()
            label_modalite.grid_forget()
            btn_add_modal.grid_forget()
            btn_del_modal.grid_forget()
            list_box_modal.grid_forget()
