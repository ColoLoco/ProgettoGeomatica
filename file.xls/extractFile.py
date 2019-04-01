#Leggere ed estrarre i dati da un file.xls
import pandas

#Apro il file xls
df = pandas.read_excel('banca.xls')

#Stampo i nomi delle colonne in una lista
print df.columns

#Stampo le righe di una determinata colonna 
values = df['nome,C,80'].values
print(values)

# stampo righe e colonne 
FORMAT = ['id,N,10,0', 'nome,C,80', 'Indirizzo,C,80','Cap,N,10,0']
df_selected = df[FORMAT]
print(df_selected)


