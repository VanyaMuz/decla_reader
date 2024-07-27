
import pymupdf  
import pandas as pd

from tkinter import filedialog
from tkinter import messagebox as mbox
from tkinter.messagebox import showinfo, askyesno, showerror
from tkinter import ttk, Tk
from tkinter import PhotoImage

import sys
import os
import logging
import threading

logging.basicConfig(level=logging.INFO, filename="pdf_reader.log", filemode='w',
                                                format="%(asctime)s %(levelname)s %(message)s")


cols = ["Страница", "№ Товара", "ТНВЭД", "Цена товара", "Основа начисления", "Ставка", "Сумма"] 
rows = [] 
df = pd.DataFrame(rows, columns=cols)
mbox.showinfo("Что делать", f"Может уже выберешь файл pdf для обработки?")


def create_df(page_num, kod_count, tnvd, cena, osnova_nach, stavka_proc, sum_df):
    df.loc[len(df)] = {'Страница': page_num+1, 
                            '№ Товара': kod_count, 
                            'ТНВЭД': tnvd,
                            'Цена товара': cena,
                            'Основа начисления': osnova_nach,
                            'Ставка': stavka_proc, 
                            'Сумма': sum_df}
    return 

def executing_data(pb, root, value_label, image, image2, image3):
    logging.info('Start executing')
    kod_count = 0
    for page_num in range(doc.page_count):
        pb['value'] = page_num+1

        value_label['text'] = f"Обработка страницы: {pb['value']} из {doc.page_count}"

        if page_num%3==0:
            value_label['image'] = image2
        else:
            value_label['image'] = image
        root.update()

        print(f'Страница', page_num+1, 'из', doc.page_count, 'в обработке...' , end='\r')
        page1 = doc[page_num]
        words = page1.get_text("words")
        strana = []
        osnova = []
        stavka = []
        cena_tovara = []
        tam_stoim =[]
        if page1.search_for('33 Код товара') != []:
            kod_tovara = page1.search_for('33 Код товара')
            strana = page1.search_for('Код стр.происх')
            osnova = page1.search_for('Основа начисления')
            stavka = page1.search_for('Ставка')
            cena_tovara = page1.search_for('42 Цена товара')
            tam_stoim = page1.search_for('45 Таможен')

            if strana ==[] or osnova ==[] or stavka ==[] or cena_tovara ==[] or tam_stoim ==[]:
                
                logging.info(f'Error in loop executing: \nstrana  {strana} \
                                                        \nosnova  {osnova}\
                                                        \nstavka {stavka}\
                                                        \ncena_tovara  {cena_tovara}\
                                                        \ntam_stoim {tam_stoim}')
            delta_y = 0
            if page1.search_for('54 Место и дата') != []:   # if it is the first page read the second line e.g. 2010
                delta_y = 8

            i = 0
            for k in kod_tovara:

                cena = ''
                tnvd = ''
                osnova_nach = ''
                stavka_proc = ''
                sum_df = ''
                vid_flag=False
                st_flag = False
                sum_flag = False

                try:
                    tnvd_coord = pymupdf.Rect(kod_tovara[i][0], kod_tovara[i][1]+1, strana[i][2], strana[i][1])
                    cena_coord = pymupdf.Rect(cena_tovara[i][0], cena_tovara[i][1]+1, tam_stoim[i][2]+47, tam_stoim[i][1])
                    vid_coord = pymupdf.Rect(osnova[i][0]-38, osnova[i][1]+12+delta_y, osnova[i][0]-10, osnova[i][3]+18+delta_y)
                    osnova_coord = pymupdf.Rect(osnova[i][0], osnova[i][1]+12+delta_y, osnova[i][2]+9, osnova[i][3]+18+delta_y)
                    stavka_coord = pymupdf.Rect(stavka[i][0], osnova[i][1]+12+delta_y, stavka[i][2]+19, osnova[i][3]+20+delta_y)
                    sum_stavka_coord = pymupdf.Rect(stavka[i][2]+25, osnova[i][1]+12+delta_y, stavka[i][2]+93, osnova[i][3]+20+delta_y)
                    
                except:
                    
                    showerror(title="Ошибка", message="Какой-то косяк, смотри логи")
                    logging.info('Error in coordinates')
                    logging.info(sys.exc_info())
                    root.destroy()
                    
                kod_count+=1
                l=0
                for w in words:
                    
                    if pymupdf.Rect(w[:4])in sum_stavka_coord:
                        if sum_flag:
                            sum_df = w[4]
                            sum_df = sum_df.replace('|', '')
                            sum_df = float(sum_df)
                            
                            sum_flag = False
                    if pymupdf.Rect(w[:4])in vid_coord:
                        if w[4] == '2010':
                            vid_flag=True
                    if pymupdf.Rect(w[:4])in stavka_coord:
                        if st_flag:
                            stavka_proc = w[4]

                            st_flag = False
                    if pymupdf.Rect(w[:4])in osnova_coord:
                        if w[4] == '│':
                            pass
                        else:
                            if vid_flag:
                                osnova_nach = w[4]
                                osnova_nach = osnova_nach.replace('|', '')
                                osnova_nach = float(osnova_nach)
                                                    
                                vid_flag=False
                                st_flag = True
                                sum_flag = True
                    if pymupdf.Rect(w[:4])in tnvd_coord:
                        tnvd = w[4]
                    
                    if pymupdf.Rect(w[:4])in cena_coord:
                        cena = float(w[4])
                        
                    if l == len(words)-1:
                        create_df(page_num, kod_count, tnvd, cena, osnova_nach, stavka_proc, sum_df)
                        logging.info(f'The line {kod_count} added to df')

                    l+=1
                i+=1
    value_label['text'] = f"Готово! Обработано страниц: {pb['value']} из {doc.page_count}"
    value_label['image'] = image3
    logging.info('End executing, decode to excel')
    df.to_excel(output_file, index=False)
    mbox.showinfo("Информация", f"Ну все, вся нужная инфа в файле {output_file}, не благодари!")
    root.destroy()

def main():
    logging.info('Start from main')
    root=Tk()
    root.geometry('%dx%d+%d+%d' % (300, 120, 500, 150))
    root.attributes("-topmost", True)
    root.protocol("WM_DELETE_WINDOW", lambda: None)
    root.resizable(0,0)
    root.title('Че то делает...')
    root.iconbitmap("_internal\myIcon.ico")
    image = PhotoImage(file="_internal\Image.png")
    image2 = PhotoImage(file="_internal\Image2.png")
    image3 = PhotoImage(file="_internal\Image3.png")
    pb = ttk.Progressbar(root, maximum= doc.page_count, orient='horizontal', mode='determinate', length=280)
    pb.grid(column=0, row=0, columnspan=2, padx=10, pady=20)
    value_label = ttk.Label(root, image=image, compound='left', text= f"Обработка страницы: {pb['value']} из {doc.page_count}")
    value_label.grid(column=0, row=1, columnspan=2)
    


    t1 = threading.Thread(target=executing_data, args=(pb, root, value_label, image, image2, image3))
    t1.daemon = True
    logging.info('Thread started')
    t1.start()
    root.mainloop()
    
if __name__ == '__main__':

    try:
        input_flag = True
        exit_flag = False
        while input_flag:
            
            ftypes = [('Эй только PDF', '*.pdf')]
            dlg = filedialog.Open(filetypes = ftypes)
            input_file = dlg.show()
            if input_file =='':
                result = askyesno(title="Ошибка", message="Файл не выбран. Выбрать файл?")
                if result: 
                    continue
                else: 
                    showinfo("Инфо", "Как хошь. Выход из программы")
                    logging.info('The file was not selected')
                    exit_flag = True
                    sys.exit()
            output_file = os.path.basename(input_file).replace('.pdf', '.xlsx')
            doc = pymupdf.open(input_file)
            logging.info(f'File {input_file} is open')
            break
        main()
        logging.info('Prog ended')
    except:
        if exit_flag==False: 
            showerror(title="Ошибка", message="Какой-то косяк, смотри логи")
        logging.error('Oops', exc_info=True)
        logging.info('Prog stopped due to an error')
        