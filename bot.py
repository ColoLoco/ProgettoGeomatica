#-*- coding: utf-8 -*-
import telepot
from time import sleep
from telepot.namedtuple import ReplyKeyboardMarkup, ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton
import pandas
def on_chat_message(msg):

	content_type, chat_type, chat_id = telepot.glance(msg)

	#Estraggo dati utenti
        username = ""
        full_name = ""
        try:
		try:
    			username = msg['chat']['username']
   		except:
        		username = "utente"
        	full_name = msg['chat']['first_name']
            	full_name += ' ' + msg['chat']['last_name']
        except KeyError:
            pass

	txt = msg['text']
	#Start
	if 'text' in msg and msg['text'] == '/start':
		markup = ReplyKeyboardMarkup(keyboard=[[("Cerca per luogo 🏔"),("Cerca per punto di interesse 🌆")]])
		bot.sendMessage(chat_id, "Benvenuto "+username+" su GeoBot!\nQui puoi visitare i luoghi e i punti di interesse che sono stati mappati!", 
			reply_markup = markup)
	#Ricerca per luogo
	elif txt.startswith("Cerca per luogo"):
		bot.sendMessage(chat_id, "Ecco i luoghi che puoi consultare:\n-->  /Acqualagna\n-->  /Furlo\n-->  /Calmazzo\n-->  /Monti??")
	#Ricerca per punto di interesse
	elif txt.startswith("Cerca per punto di interesse"):
		markup = ReplyKeyboardMarkup(keyboard=[[("Amministrazione"),("Carabinieri"),("Sanita'")],[("Servizi Culturali"),("Banche"),("Istruzione")],
			[("Sport"),("Fermate Bus"),("Svago")],[("Ristoranti"),("Negozi"),("Alberghi"),("Tartufi")]])
		bot.sendMessage(chat_id, "Ecco i punti di interrese che puoi cercare:", reply_markup=markup)
	#Acqualagna
	elif txt.startswith("/Acqualagna"):
		bot.sendMessage(chat_id,"sono su acqualagna")
	#Furlo
	elif txt.startswith("/Furlo"):
		bot.sendMessage(chat_id, "sono sul furlo")
	#Calmazzo
	elif txt.startswith("/Calmazzo"):
		bot.sendMessage(chat_id, "sono su calmazzo")
	#Monti
	elif txt.startswith("/Monti"):
		bot.sendMessage(chat_id, "sono sui monti")
	#Amministrazione
	elif txt.startswith("Amministrazione"):
		bot.sendMessage(chat_id, 'ok')
	#Carabinieri
	elif txt.startswith("Carabinieri"):
		bot.sendMessage(chat_id, 'ok')
	#Sanità
	elif txt.startswith("Sanita'"):
		bot.sendMessage(chat_id, 'ok')
	#Servizi
	elif txt.startswith("Servizi"):
		bot.sendMessage(chat_id, 'ok')
	#Banche
	elif txt.startswith("Banche"):
		result = extract_file(txt)
		bot.sendMessage(chat_id, 'Visualizzo le banche:\n')

		#vedi sito https://www.geeksforgeeks.org/reading-excel-file-using-python/
	#Istruzione
	elif txt.startswith("Istruzione"):
		bot.sendMessage(chat_id, 'ok')
	#Sport
	elif txt.startswith("Sport"):
		bot.sendMessage(chat_id, 'ok')
	#Fermate
	elif txt.startswith("Fermate"):
		bot.sendMessage(chat_id, 'ok')
	#Svago
	elif txt.startswith("Svago"):
		bot.sendMessage(chat_id, 'ok')
	#Ristoranti
	elif txt.startswith("Ristoranti"):
		bot.sendMessage(chat_id, 'ok')
	#Negozi
	elif txt.startswith("Negozi"):
		bot.sendMessage(chat_id, 'ok')
	#Alberghi
	elif txt.startswith("Alberghi"):	
		bot.sendMessage(chat_id, 'ok')
	#Tartufi
	elif txt.startswith("Tartufi"):	
		bot.sendMessage(chat_id, 'ok')

def extract_file(txt):

	#Apro il file xls
	df = pandas.read_excel('file.xls/'+txt+'.xls')
	#Seleziono la colonna
	NameColumn = ['Luogo']
	df_selected = df[NameColumn]
	print(df_selected)

	return df_selected
	

# MAIN
print("Avvio GeoBot!")

# Start working
try:
	bot = telepot.Bot(TOKEN)
    	bot.message_loop(on_chat_message)
    	while(1):
        	sleep(10)
finally:
	print("Esci")


