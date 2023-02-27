import fitz  # pip install PyMuPDF to get fitz package
from tkinter import *  # pip install tk
from tkinter import ttk
from tkinter.ttk import Progressbar
from threading import Thread
import os
import pandas as pd
import xlwings as xw
import time


# import psutil


class ShowPdf:
    img_object_li = []

    # Added zoomDPI parameter for flexibility
    def __init__(self):
        self.frame = None
        self.display_msg = None
        self.text = None

    def pdf_view(self, master, width=1200, height=600, pdf_location="", bar=True, load="after", zoom_dpi=88):

        self.frame = Frame(master, width=width, height=height, bg="white")

        scroll_y = Scrollbar(self.frame, orient="vertical")
        scroll_x = Scrollbar(self.frame, orient="horizontal")

        scroll_x.pack(fill="x", side="bottom")
        scroll_y.pack(fill="y", side="right")

        percentage_view = 0
        percentage_load = StringVar()

        if bar is True and load == "after":
            self.display_msg = Label(textvariable=percentage_load)
            self.display_msg.pack(pady=10)

            loading = Progressbar(self.frame, orient=HORIZONTAL, length=100, mode='determinate')
            loading.pack(side=TOP, fill=X)

        self.text = Text(
            self.frame,
            yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set,
            width=width, height=height
        )
        self.text.pack(side="left", expand=True)

        scroll_x.config(command=self.text.xview)
        scroll_y.config(command=self.text.yview)

        def add_img():
            percentage_divide = 0
            print(f"pdf_location: {pdf_location}")
            open_pdf = fitz.open(pdf_location)

            for page in open_pdf:
                # Use zoomDPI parameter
                # getPixmap removed from class 'Page' after v1.19 - use 'get_pixmap'
                # pix = page.getPixmap(dpi=zoomDPI)
                pix = page.get_pixmap(dpi=zoom_dpi)
                pix1 = fitz.Pixmap(pix, 0) if pix.alpha else pix
                # 'getImageData' removed from class 'Pixmap' after v1.19 - use 'tobytes'
                # img = pix1.getImageData("ppm")
                img = pix1.tobytes("ppm")
                timg = PhotoImage(data=img)
                self.img_object_li.append(timg)
                # if bar is True and load == "after":
                #     percentage_divide = percentage_divide + 1
                #     percentage_view = (float(percentage_divide)/float(len(open_pdf))*float(100))
                #     loading['value'] = percentage_view
                #     percentage_load.set(f"Please wait!, your pdf is loading {int(math.floor(percentage_view))}%")
            if bar is True and load == "after":
                loading.pack_forget()
                self.display_msg.pack_forget()

            for i in self.img_object_li:
                self.text.image_create(END, image=i)
                self.text.insert(END, "\n\n")
            self.text.configure(state="disabled")

        def start_pack():
            t1 = Thread(target=add_img)
            t1.start()

        if load == "after":
            master.after(250, start_pack)
        else:
            start_pack()
        self.frame.pack_forget()

        return self.frame


class App(ShowPdf):
    def __init__(self):
        super(App, self).__init__()
        self.col_adr_ligne = None
        self.root = Tk()
        self.root.title("App 3.0 - DIGITALISATION SDA - developed by Abdellaziz HELOUI")
        self.root.geometry("700x980")
        self.root.iconbitmap(r'LOGO_SEB.ico')
        self.root.state("zoomed")

        # Path for text file (address of station)
        self.path_user = os.path.expanduser("~")
        self.path_file_etape = self.path_user + r"\Documents\Etape.txt"

        # path for PDF files
        self.directory_pdf_path = r"\\iti13nt\Digitalisation\Document pour ligne"

        # excel file
        self.path_file_xl = self.directory_pdf_path + r"\01 DEV\Recettes_courantes.xlsx"

        # Variables
        self.recette = []
        self.produit = ""
        self.gamme = ""
        self.compteur = "Pas connecté.."
        self.counter_row = None
        self.df_excel = ""
        self.result_len_recette = False
        self.df_recette = ""
        self.lettre_ligne = self.read_text_file()[0]
        self.formule_xl_proserver = f"=PROSERVR|'Montage_Ligne_{self.lettre_ligne}.#INTERNAL'!"
        self.file_xl = f"NE PAS FERMER - ligne_{self.lettre_ligne}_poste_{self.read_text_file()[1:]}.xlsx"

        self.new_dataframe()

        # Counter
        text_label_counter = f"{time.strftime('%H:%M:%S')} | " \
                             f"LIGNE {self.lettre_ligne}, POSTE {self.read_text_file()[1:]} | " \
                             f"Compteur: {self.compteur}"
        self.label_counter = Label(
            self.root,
            text=text_label_counter,
            font=("Arial", 14, "bold"),
            fg="blue"
        )

        # List option
        self.comboBox = ttk.Combobox(
            self.root,
            font=("Arial", 14, "bold"),
            width=9,
            values=["MODOP", "SECURITE", "FLASH"],
        )
        self.comboBox.current(0)
        self.comboBox.bind("<<ComboboxSelected>>", self.action)
        print(f"sélection: {self.comboBox.get()}")

        # display pdf file
        self.d_width = 92
        self.d_height = 57
        self.d = self.pdf_view(
            self.root,
            pdf_location=f"{self.search_pdf_file(self.comboBox.get())}",
            width=self.d_width, height=self.d_height,
        )
        self.d.place(x=10, y=60)

    def action(self, event):
        # Action for choice selected (show the pdf file selected)
        selected = self.comboBox.get()
        print(f"selected: {selected}")
        self.change_pdf(self.search_pdf_file(selected))
        self.d = self.pdf_view(self.root, width=self.d_width, height=self.d_height)
        self.d.place(x=10, y=60)

    # Clear and replace the pdf
    def change_pdf(self, new_pdf_location):
        self.img_object_li = []
        self.text.configure(state="normal")
        self.text.delete("1.0", END)
        self.pdf_view(self.frame, pdf_location=new_pdf_location)

    # Read and extract the text in the text file
    def read_text_file(self):
        # Return the content file "Etape.txt"
        with open(self.path_file_etape, "r") as etape:
            return etape.read()

    # Set the variable in function of the assembly line
    def set_etape(self):
        self.formule_xl_proserver = f"=PROSERVR|'Montage_Ligne_{self.lettre_ligne}.#INTERNAL'!"

    # Create the dataframe and the Excel file
    def new_dataframe(self):
        lignes = ['Ligne_A', 'Ligne_C', 'Ligne_E', 'Ligne_F']

        variables = [
            'EUROPE', 'ADVANCE_GENIUS', 'COOKEO_1', 'COOKEO_2', 'EXTRA', 'GENIUS_2EN1', 'GENIUS_XL', 'RELIFT_V2',
            'REPRISE_ACTIFRY', 'COMPTEUR'
        ]

        # row address of variable "COMPTEUR"
        self.counter_row = len(variables) + 1

        tableau = []
        for i in variables:
            ligne = []
            for j in lignes:
                # Fills the list with 0 if the recipe is not present on the assembly line
                match j:
                    case "Ligne_A":
                        match i:
                            case "COOKEO_1":
                                ligne.append(0)
                            case "COOKEO_2":
                                ligne.append(0)
                            case "REPRISE_ACTIFRY":
                                ligne.append(0)
                            case "COMPTEUR":
                                ligne.append(f"=PROSERVR|'Montage_{j}.#INTERNAL'!CPT1.CV")
                            case _:
                                ligne.append(f"=PROSERVR|'Montage_{j}.#INTERNAL'!{i}")
                    case "Ligne_C":
                        match i:
                            case "ADVANCE_GENIUS":
                                ligne.append(0)
                            case "COOKEO_1":
                                ligne.append(0)
                            case "COOKEO_2":
                                ligne.append(0)
                            case "EXTRA":
                                ligne.append(0)
                            case "RELIFT_V2":
                                ligne.append(0)
                            case "REPRISE_ACTIFRY":
                                ligne.append(0)
                            case "COMPTEUR":
                                ligne.append(f"=PROSERVR|'Montage_{j}.#INTERNAL'!CPT1.CV")
                            case _:
                                ligne.append(f"=PROSERVR|'Montage_{j}.#INTERNAL'!{i}")
                    case "Ligne_E":
                        match i:
                            case "COOKEO_1":
                                ligne.append(0)
                            case "COOKEO_2":
                                ligne.append(0)
                            case "GENIUS_2EN1":
                                ligne.append(0)
                            case "GENIUS_XL":
                                ligne.append(0)
                            case "REPRISE_ACTIFRY":
                                ligne.append(0)
                            case "COMPTEUR":
                                ligne.append(f"=PROSERVR|'Montage_{j}.#INTERNAL'!CPT1.CV")
                            case _:
                                ligne.append(f"=PROSERVR|'Montage_{j}.#INTERNAL'!{i}")
                    case "Ligne_F":
                        match i:
                            # case "EUROPE":
                            #     ligne.append(1)
                            case "ADVANCE_GENIUS":
                                ligne.append(0)
                            case "EXTRA":
                                ligne.append(0)
                            case "GENIUS_2EN1":
                                ligne.append(0)
                            case "GENIUS_XL":
                                ligne.append(0)
                            case "RELIFT_V2":
                                ligne.append(0)
                            case "REPRISE_ACTIFRY":
                                ligne.append(0)
                            case "COMPTEUR":
                                ligne.append(f"=PROSERVR|'Montage_{j}.#INTERNAL'!CPT1.CV")
                            case _:
                                ligne.append(f"=PROSERVR|'Montage_{j}.#INTERNAL'!{i}")

            tableau.append(ligne)
        df = pd.DataFrame(tableau[0:], columns=lignes)
        # Rename index with variables name
        df.rename(index={x: y for x, y in zip(df.index, variables)}, inplace=True)

        # Create an Excel file if not exist
        if self.file_xl not in os.listdir("."):
            print("Fichier inexistant")
            with pd.ExcelWriter(self.file_xl) as writer:
                df.to_excel(
                    writer,
                    sheet_name=f'Ligne {self.lettre_ligne} Poste {self.read_text_file()[1:]}',
                    index=True
                )
                # writer.save()

    def search_pdf_file(self, option):
        self.set_recette()

        # reformatting variables to match pdf filenames
        match self.produit:
            case "GENIUS_2EN1":
                self.produit = "GENIUS_21"

            case "RELIFT_V2":
                self.produit = "RELIFTV2"

        print(f"Produit début: {self.produit}, recette: {self.recette}")
        if self.result_len_recette:
            directory = os.listdir(self.directory_pdf_path)
            for file in directory:
                match option:
                    case "MODOP":
                        if self.produit in file:
                            if f"ETAPE{self.read_text_file()[1:]}" in file:
                                return self.directory_pdf_path + f"\\{file}"
                    case "SECURITE":
                        if "Securite_MONTAGE_STANDARD" in file:
                            return self.directory_pdf_path + f"\\{file}"
                    case "FLASH":
                        # if self.produit in file:
                        if option in file:
                            if f"LIGNE_{self.lettre_ligne}" in file:
                                if f"POSTE_{self.read_text_file()[1:]}" in file:
                                    print(f"File searched: {file}")
                                    return self.directory_pdf_path + f"\\{file}"

        else:
            return self.directory_pdf_path + "\\0_Repos.pdf"

    def set_recette(self):
        # Initialization
        self.recette = []
        adr = None

        # Waiting for Excel file
        while True:
            try:
                wb = xw.Book(self.file_xl)
                sheet = wb.sheets[0]
                sheet.autofit()
                break
            except Exception:
                time.sleep(2)
        # print("Le fichier a été ouvert avec succès")

        # Return the address value
        for i in range(0, sheet.used_range.last_cell.row):
            for j in range(1, sheet.used_range.last_cell.column):
                cel = sheet[i, j]
                if cel.value == f'Ligne_{self.lettre_ligne}':
                    adr = cel.address
                    # print(f"Adresse: {adr}")
                    break

        plage = sheet.range(f'{adr[1]}2:{adr[1]}' + f'{self.counter_row - 1}')
        for cellule in plage:
            if cellule.value == 1:
                self.recette.append(cellule.row)
        # print(self.lettre_ligne, recette)

        # Return the name in column "A" (Variables)
        if len(self.recette) > 1:
            self.gamme = sheet[f"A{self.recette[0]}"].value
            self.produit = sheet[f"A{self.recette[1]}"].value
            # Vérifier si la recette courante a été modifiée et actualise les infos
            if self.produit != sheet[f"A{self.recette[1]}"].value:
                self.produit = sheet[f"A{self.recette[1]}"].value
                # TODO place here the code to update the pdf showing
                self.action(self.comboBox.get())
                print("Product changed...")
            self.result_len_recette = True
            # print(f"Gamme: {self.gamme}, Produit: {self.produit}")
        else:
            self.result_len_recette = False
            # print("Ligne en repos...")

        self.col_adr_ligne = adr[1]

        if sheet[f"{adr[1]}{self.counter_row}"].value is None:
            self.compteur = "pas connecté..."
        else:
            self.compteur = int(sheet[f"{adr[1]}{self.counter_row}"].value)
        # print(f"Compteur ligne {self.lettre_ligne}: {compteur}")

        del adr
        del plage
        del cel
        del sheet
        del wb

    def update_label(self):
        current_time = time.strftime("%H:%M:%S")

        wb = xw.Book(self.file_xl)
        sheet = wb.sheets[0]

        if sheet[f"{self.col_adr_ligne}{self.counter_row}"].value is not None:
            self.compteur = int(sheet[f"{self.col_adr_ligne}{self.counter_row}"].value)
        else:
            self.compteur = "Pas connecté..."

        self.label_counter["text"] = f"{current_time} | " \
                                     f"LIGNE {self.lettre_ligne}, " \
                                     f"POSTE {self.read_text_file()[1:]}  " \
                                     f"{self.produit} | " \
                                     f"Compteur : {self.compteur}"

        del current_time
        del self.compteur
        del wb
        del sheet

        self.root.after(500, self.update_label)

    def main(self):
        print(f"Main | heure lancement : {time.strftime('%H:%M:%S')}")
        self.comboBox.place(x=10, y=15)
        self.label_counter.place(x=130, y=15)

    def run(self):
        self.main()
        self.new_dataframe()
        self.root.after(20, self.update_label)
        self.root.mainloop()


if __name__ == '__main__':
    App().run()
