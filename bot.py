#-*- coding: utf-8 -*-
import telepot
from time import sleep
from telepot.namedtuple import ReplyKeyboardMarkup, ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton
import xlrd

#Dizionari
user_state = {}
bank = {}
count = {}
nameFile = {}

#Banche             OK
#Fermate            OK
#Assicurazioni      OK
#Ristoranti         OK
#Tartufi            OK
#Sport              OK
#Musei              OK
#Svago              OK
#Negozi             OK

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
		markup = ReplyKeyboardMarkup(keyboard=[[("Amministrazione"),("Sicurezza"),("Sanita'")],[("Musei"),("Banche"),("Istruzione")],
			[("Sport"),("Fermate Bus"),("Svago")],[("Negozi"),("Alberghi"),("Tartufi")],[("Assicurazioni"),("Ristoranti")]])
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
	#Musei
	elif txt.startswith("Musei"):
	        initialCategory("Musei", 11, txt, chat_id)
        #Banche
	elif txt.startswith("Banche") or txt.startswith("Assicurazioni"):
                if txt.startswith("Banche"):
                    initialCategory("Banche", 1, txt, chat_id)
                else:
                    initialCategory("Assicurazioni", 1, txt, chat_id)
                    #vedi sito https://www.geeksforgeeks.org/reading-excel-file-using-python/
	#Istruzione
	elif txt.startswith("Istruzione"):
		bot.sendMessage(chat_id, 'ok')
	#Sport
	elif txt.startswith("Sport"):
		initialCategory("Sport", 9, txt, chat_id)
        #Fermate
	elif txt.startswith("Fermate"):
		initialCategory("Fermate Bus", 3, txt, chat_id)
	#Svago
	elif txt.startswith("Svago"):
		initialCategory("Svago", 13, txt, chat_id)
	#Ristoranti
	elif txt.startswith("Ristoranti"):
		initialCategory("Ristoranti", 5, txt, chat_id)
	#Negozi
	elif txt.startswith("Negozi"):
		initialCategory("Negozi", 15, txt, chat_id)
	#Alberghi
	elif txt.startswith("Alberghi"):	
		initialCategory("Alberghi", 17, txt, chat_id)
	#Tartufi
	elif txt.startswith("Tartufi"):	
		initialCategory("Tartufi", 7, txt, chat_id)
        #1 - Banche
        elif user_state[chat_id] == 1:
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Orario"),("Note"),("Foto"),("ATM")],[("Telefono"),("Sito"),('Home')]])
            keyboardCategory(txt, markup, 2, chat_id)
        #2 - Banche            
        elif user_state[chat_id] == 2:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Orario"):
                 bot.sendMessage(chat_id, "Orario: "+ sheet.cell_value(count[chat_id],7))
                        
            elif txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Note: "+sheet.cell_value(count[chat_id],6)+"\n")
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],5)+"\n")
            elif txt.startswith("ATM"):
                 result = sheet.cell_value(count[chat_id],2)
                 if result == 0:
                     bot.sendMessage(chat_id, "Non è presente")
                 else:
                     bot.sendMessage(chat_id, "E' presente")
            elif txt.startswith("Sito"):
                keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Sito web",url=sheet.cell_value(count[chat_id],4))],])
                bot.sendMessage(chat_id, "Ecco il sito:"+"\n", reply_markup=keyboard)
            elif txt.startswith("Telefono"):
                bot.sendMessage(chat_id, "Telefono: "+str(sheet.cell_value(count[chat_id],3)))
            elif txt.startswith("Home"):
                home(chat_id,username)
        #3 - Fermate            
        elif user_state[chat_id] == 3:
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Riparata"),("Sito")],[("Note"),("Home")]])
            keyboardCategory(txt, markup, 4, chat_id)
        #4 - Fermate            
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
            elif txt.startswith("Home"):
                home(chat_id,username)
        #5 - Ristoranti            
        elif user_state[chat_id] == 5:
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Specialita"),("Orari"),("Telefono")],[("Sito"),("Note"),("Foto"),("Home")]])
            keyboardCategory(txt, markup, 6, chat_id) 
        #6 - Ristoranti            
        elif user_state[chat_id] == 6:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Info: "+sheet.cell_value(count[chat_id],7)+"\n")
            elif txt.startswith("Orari"):
                 bot.sendMessage(chat_id, "Orari: "+sheet.cell_value(count[chat_id],3)+"\n")
            elif txt.startswith("Specialita"):
                bot.sendMessage(chat_id, "Specialita': "+deEmojify(sheet.cell_value(count[chat_id],2)+"\n"))
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],6)+"\n")
            elif txt.startswith("Sito"):
                try:
                    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Sito web",url=sheet.cell_value(count[chat_id],5))],])
                    bot.sendMessage(chat_id, "Ecco il sito:"+"\n", reply_markup=keyboard)
                except:
                    bot.sendMessage(chat_id, "Non è presente il sito!")
            elif txt.startswith("Telefono"):
                bot.sendMessage(chat_id, "Telefono: "+str(sheet.cell_value(count[chat_id],4)))    
            elif txt.startswith("Home"):
                home(chat_id,username)
        #7 - Tartufi           
        elif user_state[chat_id] == 7:
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Orario"),("Note"),("Foto")],[("Telefono"),("Sito"),('Home')]])
            keyboardCategory(txt, markup, 8, chat_id)
        #8 - Tartufi            
        elif user_state[chat_id] == 8:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Orario"):
                 bot.sendMessage(chat_id, "Orario: "+ sheet.cell_value(count[chat_id],6))
            elif txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Note: "+sheet.cell_value(count[chat_id],5)+"\n")
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],4)+"\n")
            elif txt.startswith("Sito"):
                try:
                    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Sito web",url=sheet.cell_value(count[chat_id],3))],])
                    bot.sendMessage(chat_id, "Ecco il sito:"+"\n", reply_markup=keyboard)
                except:
                    bot.sendMessage(chat_id, "Non è presente il sito!")
            elif txt.startswith("Telefono"):
                bot.sendMessage(chat_id, "Telefono: "+str(sheet.cell_value(count[chat_id],2)))
            elif txt.startswith("Home"):
                home(chat_id,username)
        #9 - Sport            
        elif user_state[chat_id] == 9:
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Riparato"),("Tipologia")],[("Foto"),("Note"),("Home")]])
            keyboardCategory(txt, markup, 10, chat_id)
        #10 - Sport            
        elif user_state[chat_id] == 10:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Note: "+sheet.cell_value(count[chat_id],5)+"\n")
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],4)+"\n")
            elif txt.startswith("Riparato"):
                 result = sheet.cell_value(count[chat_id],3)
                 if result == 0:
                     bot.sendMessage(chat_id, "Non è presente")
                 else:
                     bot.sendMessage(chat_id, "E' presente")
            elif txt.startswith("Tipologia"):
                 bot.sendMessage(chat_id, "Tipologia: "+sheet.cell_value(count[chat_id],2)+"\n")
            elif txt.startswith("Home"):
                home(chat_id,username)
        #11 - Musei           
        elif user_state[chat_id] == 11:
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Orari"),("Telefono"),("Sito")],[("Foto"),("Note"),("Home")]])
            keyboardCategory(txt, markup, 12, chat_id)
        #12 - Musei            
        elif user_state[chat_id] == 12:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Note: "+sheet.cell_value(count[chat_id],6)+"\n")
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],5)+"\n")
            elif txt.startswith("Orari"):
                  bot.sendMessage(chat_id, "Orario: "+sheet.cell_value(count[chat_id],2)+"\n")           
            elif txt.startswith("Telefono"):
                 bot.sendMessage(chat_id, "Telefono: "+sheet.cell_value(count[chat_id],3)+"\n")
            elif txt.startswith("Sito"):
                try:
                    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Sito web",url=sheet.cell_value(count[chat_id],4))],])
                    bot.sendMessage(chat_id, "Ecco il sito:"+"\n", reply_markup=keyboard)
                except:
                    bot.sendMessage(chat_id, "Non è presente il sito!")    
            elif txt.startswith("Home"):
                home(chat_id,username)   
        #13 - Svago         
        elif user_state[chat_id] == 13:
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Tipologia"),("Riparato"),("Foto")],[("Note"),("Home")]])
            keyboardCategory(txt, markup, 14, chat_id)
        #14 - Svago            
        elif user_state[chat_id] == 14:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Note: "+sheet.cell_value(count[chat_id],5)+"\n")
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],4)+"\n")
            elif txt.startswith("Riparato"):
                 result = sheet.cell_value(count[chat_id],3)
                 if result == 0:
                     bot.sendMessage(chat_id, "Non è presente")
                 else:
                     bot.sendMessage(chat_id, "E' presente")
            elif txt.startswith("Tipologia"):
                 bot.sendMessage(chat_id, "Tipologia: "+sheet.cell_value(count[chat_id],2)+"\n")
            elif txt.startswith("Home"):
                home(chat_id,username)
        #15 - Negozi            
        elif user_state[chat_id] == 15:
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Tipologia"),("Orari"),("Foto"),("Sito")],[("Note"),("Telefono"),("Home")]])
            keyboardCategory(txt, markup, 16, chat_id)
        #16 - Negozi           
        elif user_state[chat_id] == 16:
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Note: "+sheet.cell_value(count[chat_id],7)+"\n")
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],6)+"\n")
            elif txt.startswith("Orari"):
                  bot.sendMessage(chat_id, "Orario: "+sheet.cell_value(count[chat_id],3)+"\n")           
            elif txt.startswith("Telefono"):
                 bot.sendMessage(chat_id, "Telefono: "+sheet.cell_value(count[chat_id],4)+"\n")
            elif txt.startswith("Tipologia"):
                 bot.sendMessage(chat_id, "Tipologia: "+sheet.cell_value(count[chat_id],2)+"\n")
            elif txt.startswith("Sito"):
                try:
                    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Sito web",url=sheet.cell_value(count[chat_id],5))],])
                    bot.sendMessage(chat_id, "Ecco il sito:"+"\n", reply_markup=keyboard)
                except:
                    bot.sendMessage(chat_id, "Non è presente il sito!")    
            elif txt.startswith("Home"):
                home(chat_id,username) 
        #17 - Alberghi           
        elif user_state[chat_id] == 17:  
            txt = msg['text']
            markup = ReplyKeyboardMarkup(keyboard=[[("Telefono"),("Foto")],[("Note"),("Sito"),("Home")]])
            keyboardCategory(txt, markup, 18, chat_id)
        #18 - Alberghi           
        elif user_state[chat_id] == 18:
            txt = msg['text']
            wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
            sheet = wb.sheet_by_index(0) 
            #estraggo le informazioni
            if txt.startswith("Note"):
                 bot.sendMessage(chat_id, "Note: "+sheet.cell_value(count[chat_id],5)+"\n")
            elif txt.startswith("Foto"):
                 bot.sendMessage(chat_id, "Foto: "+sheet.cell_value(count[chat_id],4)+"\n")
            elif txt.startswith("Telefono"):
                 bot.sendMessage(chat_id, "Telefono: "+sheet.cell_value(count[chat_id],2)+"\n")
            elif txt.startswith("Sito"):
                try:
                    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Sito web",url=sheet.cell_value(count[chat_id],3))],])
                    bot.sendMessage(chat_id, "Ecco il sito:"+"\n", reply_markup=keyboard)
                except:
                    bot.sendMessage(chat_id, "Non è presente il sito!")    
            elif txt.startswith("Home"):
                home(chat_id,username)

            




# estrae i dati file xls  
def extract_file(txt):

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

# stampa la lista della categoria scelta
def initialCategory(name, user, txt, chat_id):
    result = extract_file(txt)
    bot.sendMessage(chat_id, '*Digita il numero per avere maggiori informazioni.*\n'+result, parse_mode="Markdown")
    user_state[chat_id] = user
    nameFile[chat_id] = name

#stampa i pulsanti specifici per il tipo di locale/servizio per la categoria scelta
def keyboardCategory(txt, markup, user, chat_id):
    wb = xlrd.open_workbook('file.xls/'+nameFile[chat_id]+'.xls')
    sheet = wb.sheet_by_index(0) 
    if txt.isnumeric() and int(txt) <= sheet.nrows-1 and  int(txt) > 0:
          count[chat_id] = int(txt)
          bot.sendMessage(chat_id, 'Ecco cosa puoi visualizzare:', reply_markup=markup)
          user_state[chat_id] = user
    else:
          bot.sendMessage(chat_id, 'Formato errato!')

# si ritorna alla home
def home(chat_id,username):
    markup = ReplyKeyboardMarkup(keyboard=[[("Cerca per luogo 🏔"),("Cerca per punto di interesse 🌆")]])
    bot.sendMessage(chat_id, "Benvenuto "+username+" su GeoBot!\nQui puoi visitare i luoghi e i punti di interesse che sono stati mappati!", 
    		 reply_markup = markup)
#pulisce la stringa da caratteri non ascii
def deEmojify(inputString):
    return inputString.encode('ascii', 'ignore').decode('ascii')    
	
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


