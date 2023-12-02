from img2table.document import PDF
from img2table.ocr import TesseractOCR
from pypdf import PdfReader
import pandas as pd
import numpy as np
import os
from alive_progress import alive_bar

def main():
    file = "/upload/uploaded_file.pdf"
    reader = PdfReader(file)
    n=len(reader.pages)
    print("initialisation...")
    ocr = TesseractOCR()
    print()
    end = False
    c=0
    print("=> extraction des tableaux")
    with alive_bar(n) as bar:
        while not end:
            try:
                pdf = PDF(file,pages=[c])
                pdf_tables = pdf.extract_tables(ocr=ocr)
                pdf.to_xlsx(f'tables/tableau_{c+1}.xlsx',ocr=ocr)
                c=c+1
                bar()
            except:
                end = True

    print(f"{c} pages traitées")
    print("")

    def detecter_et_supprimer_excel_vide(chemin_fichier):
        df = pd.read_excel(chemin_fichier)
        if df.shape == (0,0):
            os.remove(chemin_fichier)
            return

    def parcourir_dossier_et_supprimer_fichiers_vides(chemin_dossier):
        for nom_fichier in os.listdir(chemin_dossier):
            chemin_fichier = os.path.join(chemin_dossier, nom_fichier)
            if os.path.isfile(chemin_fichier) and nom_fichier.endswith('.xlsx'):
                detecter_et_supprimer_excel_vide(chemin_fichier)

    chemin_dossier = '/tables'
    parcourir_dossier_et_supprimer_fichiers_vides(chemin_dossier)

    print("=> création du fichier")
    def parcourir_dossier_et_supprimer_fichiers_vides(chemin_dossier):
        L = []
        for nom_fichier in os.listdir(chemin_dossier):
            fichier = nom_fichier.split("_")[1].split(".")[0]
            L.append(int(fichier))
        L = sorted(L)
        max_width = 0
        sheet = []
        for i in range(len(L)):
            dataframes = pd.read_excel(f"{chemin_dossier}/tableau_{str(L[i])}.xlsx", sheet_name=None)
            for sheet_name, df in dataframes.items():
                if df.shape[1] > max_width:
                    max_width = df.shape[1]
                sheet.append(sheet_name)
                #print(f"{sheet_name} / {df.shape}")
        return(L,max_width,sheet)

    chemin_dossier = '/tables'
    L , max_width, sheet= parcourir_dossier_et_supprimer_fichiers_vides(chemin_dossier)

    def file_p(sheet):
        n = sheet.split(" ")[1]
        return(f"/tables/tableau_{n}.xlsx")


    def merge(file1,file2,sheet1,sheet2,total_col):
        dataframes = pd.read_excel(file1, sheet_name=None)
        dataframes2 = pd.read_excel(file2, sheet_name=None)
        array1 = np.vstack([dataframes[sheet1].columns, dataframes[sheet1].to_numpy()])
        array2 = np.vstack([dataframes2[sheet2].columns, dataframes2[sheet2].to_numpy()])
        #total_col = max(array1.shape[1], array2.shape[1])
        array1 = np.hstack([array1, np.full((array1.shape[0], total_col - array1.shape[1]), np.nan)])
        array2 = np.hstack([array2, np.full((array2.shape[0], total_col - array2.shape[1]), np.nan)])
        blank = [""]*total_col
        result_array = np.vstack([array1, blank])
        result_array = np.vstack([result_array, array2])
        df = pd.DataFrame(result_array)
        return df

    def merge2(file1,file2,sheet2,total_col):
        dataframes  = pd.read_excel(file1, sheet_name=None)
        dataframes2 = pd.read_excel(file2, sheet_name=None)
        array1 = np.vstack([dataframes['Sheet1'].columns, dataframes['Sheet1'].to_numpy()])
        array2 = np.vstack([dataframes2[sheet2].columns, dataframes2[sheet2].to_numpy()])
        #total_col = max(array1.shape[1], array2.shape[1])
        array1 = np.hstack([array1, np.full((array1.shape[0], total_col - array1.shape[1]), np.nan)])
        array2 = np.hstack([array2, np.full((array2.shape[0], total_col - array2.shape[1]), np.nan)])
        blank  = [""]*total_col
        result_array = np.vstack([array1, blank])
        result_array = np.vstack([result_array, array2])
        df = pd.DataFrame(result_array)
        return df

    df = merge(file_p(sheet[0]),file_p(sheet[1]),sheet[0],sheet[1],max_width)
    temp_file_path = '/output_temp.xlsx'
    df.to_excel(temp_file_path, index=False)

    with alive_bar(len(sheet)-1) as bar:
        for i in range(1,len(sheet)):
            bar()
            df = merge2('/output_temp.xlsx',file_p(sheet[i]),sheet[i],max_width)
            df.to_excel(temp_file_path, index=False)

    df = df.iloc[len(sheet)-1:]

    def vider_unamed(cellule):
        return '' if 'Unnamed' in str(cellule) else cellule

    df = df.applymap(vider_unamed)

    output_file_path = '/pdf.xlsx'
    df.to_excel(output_file_path, index=False)
    print("=> fichier créé")

    fichiers_dans_dossier = os.listdir("/tables/")
    for fichier in fichiers_dans_dossier:
        chemin_fichier = os.path.join("/tables/", fichier)
        if os.path.isfile(chemin_fichier):
            os.remove(chemin_fichier)
    os.remove('/output_temp.xlsx')
    #os.remove(str(os.listdir("/home/benjamin/flask pdf2xlsx/upload/")[0]))