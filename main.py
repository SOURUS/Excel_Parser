# -*- coding: cp1251 -*-
from openpyxl import *
import xlwt

#переводим инфу из экселя в массивы
def HochladenInfo(files):
    data = {}
    for name in files:
        data[name]=[]
        wb = load_workbook(filename = name, use_iterators = True)
        ws = wb.worksheets[0]
        rowcount = 0
        for row in ws.iter_rows():
            data[name].append([])    
            for cell in row:
                data[name][rowcount].append(cell.value)
            rowcount +=1
        print 'File "'+name+'" ist hochgeladen'
        
    return data

#парсим страницы
def Parsing_usw(data):
    for a_rownum in range(len(data['AddOrEddit.xlsx'])):
        a_cell = data['AddOrEddit.xlsx'][a_rownum][0]
        MnemonikaIsFound = False
        
        for m_rownum in range(len(data['mnemonics.xlsx'])):
            m_cell = data['mnemonics.xlsx'][m_rownum][1]
            if (a_cell==m_cell): # мнемоника из перевода совпадает с то которая в бд
                MnemonikaIsFound = True
                Ubersetzung(data['mnemonics.xlsx'][m_rownum][0], a_rownum, data)
                break

        if ((MnemonikaIsFound == False) & (a_cell!=None)):
            AddMnemonik(a_cell, data)
            Ubersetzung(data['mnemonics.xlsx'][-1][0], a_rownum, data)
            break

#добавление новой мнемоники
def AddMnemonik(name, data):
    last_pos = data['mnemonics.xlsx'][-1][0] + 1
    data['mnemonics.xlsx'].append([])
    data['mnemonics.xlsx'][-1].append(last_pos)#добавляем ID к мнемонике
    data['mnemonics.xlsx'][-1].append(name)#имя
    data['mnemonics.xlsx'][-1].append(1)#Area
    #print "----------\nДобавлена новая мнемоника:", name," | с id -> ", data['mnemonics.xlsx'][-1][0],"\n----------"
            
#добавления новых переводов        
def Hinzufugen(languages, id_m, a_rownum, data):
    while len(languages)>0:
        last_pos = data['translation.xlsx'][-1][0] + 1
        data['translation.xlsx'].append([])
        data['translation.xlsx'][-1].append(last_pos)#добавляем ID к переводу
        data['translation.xlsx'][-1].append(id_m)#добавляем ID мнемноики
        data['translation.xlsx'][-1].append(languages[0])#код языка
        data['translation.xlsx'][-1].append(data['AddOrEddit.xlsx'][a_rownum][languages[0]+1])#перевод

        #print "Добавлен новый перевод для мнемоники с id ->",id_m,"| на языке", languages[0], "| Ubersetzung: ",data['translation.xlsx'][-1][3]
        del languages[0]
        
#поиск существующих и их исправление (если нужно)
def Ubersetzung(id_m, a_rownum, data):
    languages = [0,1,2] #0 - рус, 1- анг, 2-турк
    while languages: # пока не будет 3 переводов
        for t_rownum in range(len(data['translation.xlsx'])):
            x=data['translation.xlsx'][t_rownum][2]
            if x > 2: # значит это не перевод (а коммент. или транслит)
                continue
            #ниже проверяется, перевод ли это (а не транслит), есть ли такой перевод в AddOrEdit, если есть то замеянем.
            if ((id_m==data['translation.xlsx'][t_rownum][1])  and (data['AddOrEddit.xlsx'][a_rownum][x+1] != None) ):
                data['translation.xlsx'][t_rownum][3] = data['AddOrEddit.xlsx'][a_rownum][x+1] # плюс один потому что, мнемоника на первом в AddOrEddit........может стоит всеж создать классы :(
                #print "Обновили перевод дл мнемоники с id -> ",id_m,"| На языке= ",x
                languages.remove(x) # перевод найден/исправен, удаляем.
                continue
            
        Hinzufugen(languages, id_m, a_rownum, data)
        
#Сохраняем в отдельный excel
def Speichern(data, name=[]):
    counter = 0
    for doc in name:
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("MyData")
        for row_index, row in enumerate(data[name[counter]]):
            for col_index, cell_value in enumerate(row):
                sheet.write(row_index, col_index, cell_value)
                #print(row_index, col_index, cell_value)
        workbook.save("neu_"+name[counter][:-1])
        print "neu_"+name[counter], " <- ist saved"
        counter+=1
        
        
def main():
    dic = {}
    dic = HochladenInfo(['mnemonics.xlsx', 'translation.xlsx', 'AddOrEddit.xlsx'])
    print "\nThe documents were downloaded!"
    Parsing_usw(dic)
    print "\nThe documents is ready!\n"
    Speichern(dic, ['mnemonics.xlsx','translation.xlsx'])
    print "\nThe Job is dont!\n"
    raw_input("Press the enter key to exit.")
    print "Have a nice day!"
    exit()
    
if __name__ == "__main__":
    main()
