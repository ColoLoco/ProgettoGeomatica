#-*- coding: utf-8 -*-
import telepot
from time import sleep
from telepot.namedtuple import ReplyKeyboardMarkup, ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton
import xlrd

#Dizionari
nameFile = {}
user_state = {}
bank = {}
count = {}

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

        try:
            bank[chat_id]=bank[chat_id]
            user_state[chat_id]=user_state[chat_id]
            count[chat_id] = count[chat_id]
            nameFile[chat_id] = nameFile[chat_id]

        except:
            bank[chat_id] = 0
            user_state[chat_id]=0
            count[chat_id] = 0
            nameFile[chat_id] = 0

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
		markup = ReplyKeyboardMarkup(keyboard=[[("Amministrazione"),("Sicurezza"),("Sanita'")],[("Servizi Culturali"),("Banche"),("Istruzione")],
			[("Sport"),("Fermate Bus"),("Svago")],[("Negozi"),("Alberghi"),("Tartufi")],[("Assicurazione"),("Ristoranti")]])
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
	elif txt.startswith("Sicurezza"):
		bot.sendMessage(chat_id, 'ok')
	#Sanità
	elif txt.startswith("Sanita'"):
		bot.sendMessage(chat_id, 'ok')
	#Servizi
	elif txt.startswith("Servizi"):
		bot.sendMessage(chat_id, 'ok')
	#Banche
	elif txt.startswith("Banche") or txt.startswith("Assicurazione"):
		result = extract_fileBanche(txt)
                bot.sendMessage(chat_id, '*Digita il numero per avere maggiori informazioni.*\n'+result, parse_mode="Markdown")
                
                user_state[chat_id] = 1
                nameFile = "Banche"
		#vedi sito https://www.geeksforgeeks.org/reading-excel-file-using-python/
	#Istruzione
	elif txt.startswith("Istruzione"):
		bot.sendMessage(chat_id, 'ok')
	#Sport
	elif txt.startswith("Sport"):
		bot.sendMessage(chat_id, 'ok')
	#Fermate
	elif txt.startswith("Fermate"):
		result = extract_fileBanche(txt)
                bot.sendMessage(chat_id, '*Digita il numero per avere maggiori informazioni.*\n'+result, parse_mode="Markdown")
                
                nameFile[chat_id] = "Fermate Bus"
                user_state[chat_id] = 3
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
        #1 - Banche
        elif user_state[chat_id] == 1:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 

            if txt.isnumeric() and int(txt) <= sheet.nrows-1 and  int(txt) > 0:
                  count[chat_id] = int(txt)
                  markup = ReplyKeyboardMarkup(keyboard=[[("Orario"),("Note"),("Foto"),("ATM")],[("Telefono"),("Sito")]])
                  bot.sendMessage(chat_id, 'Ecco cosa puoi visualizzare:', reply_markup=markup)
                  user_state[chat_id] = 2
            else:
                  bot.sendMessage(chat_id, 'Formato errato!')
        #2 - Banche            
        elif user_state[chat_id] == 2:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Orario"):
                bot.sendMessage(chat_id, "Orario: "+str(sheet.cell_value(count[chat_id],2)))
            elif txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Note: "+sheet.cell_value(count[chat_id],7)+"\n")
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],6)+"\n")
            elif txt.startswith("ATM"):
                 result = sheet.cell_value(count[chat_id],3)
                 if result == 0:
                     bot.sendMessage(chat_id, "Non è presente")
                 else:
                     bot.sendMessage(chat_id, "E' presente")
            elif txt.startswith("Sito"):
                keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Sito web",url=sheet.cell_value(count[chat_id],5))],])
                bot.sendMessage(chat_id, "Ecco il sito:"+"\n", reply_markup=keyboard)

            elif txt.startswith("Telefono"):
                bot.sendMessage(chat_id, "Telefono: "+str(sheet.cell_value(count[chat_id],4)))
            elif txt.startswith("Indietro"):
                print('bho') 
        #3 - Fermate            
        elif user_state[chat_id] == 3:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 

            if txt.isnumeric() and int(txt) <= sheet.nrows-1 and  int(txt) > 0:
                  count[chat_id] = int(txt)
                  markup = ReplyKeyboardMarkup(keyboard=[[("Riparata"),("Sito"),("Note")]])
                  bot.sendMessage(chat_id, 'Ecco cosa puoi visualizzare:', reply_markup=markup)
                  user_state[chat_id] = 4
            else:
                  bot.sendMessage(chat_id, 'Formato errato!')
        #3 - Fermate            
        elif user_state[chat_id] == 4:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Info: "+sheet.cell_value(count[chat_id],4)+"\n")
            elif txt.startswith("Riparata"):
                 result = sheet.cell_value(count[chat_id],2)
                 if result == 0:
                     bot.sendMessage(chat_id, "Non è presente")
                 else:
                     bot.sendMessage(chat_id, "E' presente")
            elif txt.startswith("Sito"):
                keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Sito web",url=sheet.cell_value(count[chat_id],3))],])
                bot.sendMessage(chat_id, "Ecco il sito:"+"\n", reply_markup=keyboard)
            elif txt.startswith("Indietro"):
                print('bho') 

          



def extract_fileBanche(txt):

        # Give the location of the file  
        wb = xlrd.open_workbook('file.xls/'+txt+'.xls') 
        txt = ""
        sheet = wb.sheet_by_index(0) 

        count = 1
        for i in range(0,sheet.nrows - 1):
            txt += str(count)+") "
            txt += str(sheet.cell_value(count,1)+"\n")
            count += 1  

	return txt
	
# MAIN
print("Avvio GeoBot!")

# Start working
try:
	bot = telepot.Bot('584122851:AAGoR1fv83GnMYuz87TcRVYxiaWMKE_YFtA')
    	bot.message_loop(on_chat_message)
    	while(1):
        	sleep(10)
finally:
	print("Esci")


