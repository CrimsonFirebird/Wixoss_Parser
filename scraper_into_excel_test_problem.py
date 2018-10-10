from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
from enum import Enum
import xlsxwriter

#Excel Stuff Creator
workbook = xlsxwriter.Workbook('WXtest.xlsx')
worksheet_Lrig = workbook.add_worksheet('Lrig')
worksheet_Arts = workbook.add_worksheet('Arts')
worksheet_Signi = workbook.add_worksheet('Signi')
worksheet_Spell = workbook.add_worksheet('Spell')
worksheet_Key = workbook.add_worksheet('Key')
worksheet_Resona = workbook.add_worksheet('Resona')
#End of Excel file creator

global Eng
Eng = 0

#Writing of the Title for individual sheets (LRIG)
worksheet_Lrig.write(0, 0, 'Card no')
worksheet_Lrig.write(0, 1, 'Name')
worksheet_Lrig.write(0, 2, 'Rarity')
worksheet_Lrig.write(0, 3, 'Link Suffix')
worksheet_Lrig.write(0, 4, 'Effects')
worksheet_Lrig.write(0, 5, 'Kana')
worksheet_Lrig.write(0, 6, 'Romaji')
worksheet_Lrig.write(0, 7, 'Color')
worksheet_Lrig.write(0, 8, 'Card Type')
worksheet_Lrig.write(0, 9, 'Level')
worksheet_Lrig.write(0, 10, 'Limit')
worksheet_Lrig.write(0, 11, 'Grow Cost')
worksheet_Lrig.write(0, 12, 'Lrig Type')
worksheet_Lrig.write(0, 13, 'Coin')
# Start from the first cell. Rows and columns are zero indexed.
global row_Lrig
global col_Lrig
row_Lrig = 1
#End of Writing of the Title for individual sheets (LRIG)

#Writing of the Title for individual sheets (SIGNI)
worksheet_Signi.write(0, 0, 'Card no')
worksheet_Signi.write(0, 1, 'Name')
worksheet_Signi.write(0, 2, 'Rarity')
worksheet_Signi.write(0, 3, 'Link Suffix')
worksheet_Signi.write(0, 4, 'Effects')
worksheet_Signi.write(0, 5, 'Kana')
worksheet_Signi.write(0, 6, 'Romaji')
worksheet_Signi.write(0, 7, 'Color')
worksheet_Signi.write(0, 8, 'Card Type')
worksheet_Signi.write(0, 9, 'Level')
worksheet_Signi.write(0, 10, 'Power')
worksheet_Signi.write(0, 11, 'Limiting Condition')
worksheet_Signi.write(0, 12, 'Class')
# Start from the first cell. Rows and columns are zero indexed.
global row_Signi
global col_Signi
row_Signi = 1
#End of Writing of the Title for individual sheets (SIGNI)

#Writing of the Title for individual sheets (ARTS)
worksheet_Arts.write(0, 0, 'Card no')
worksheet_Arts.write(0, 1, 'Name')
worksheet_Arts.write(0, 2, 'Rarity')
worksheet_Arts.write(0, 3, 'Link Suffix')
worksheet_Arts.write(0, 4, 'Effects')
worksheet_Arts.write(0, 5, 'Kana')
worksheet_Arts.write(0, 6, 'Romaji')
worksheet_Arts.write(0, 7, 'Color')
worksheet_Arts.write(0, 8, 'Card Type')
worksheet_Arts.write(0, 9, 'Cost')
worksheet_Arts.write(0, 10, 'Limiting Condition')
worksheet_Arts.write(0, 11, 'Use Timing')
# Start from the first cell. Rows and columns are zero indexed.
global row_Arts
global col_Arts
row_Arts = 1
#End of Writing of the Title for individual sheets (ARTS)

#Writing of the Title for individual sheets (KEYS)
worksheet_Key.write(0, 0, 'Card no')
worksheet_Key.write(0, 1, 'Name')
worksheet_Key.write(0, 2, 'Rarity')
worksheet_Key.write(0, 3, 'Link Suffix')
worksheet_Key.write(0, 4, 'Effects')
worksheet_Key.write(0, 5, 'Kana')
worksheet_Key.write(0, 6, 'Romaji')
worksheet_Key.write(0, 7, 'Color')
worksheet_Key.write(0, 8, 'Card Type')
worksheet_Key.write(0, 9, 'Cost')
worksheet_Key.write(0, 10, 'Key Selection Legal?')
# Start from the first cell. Rows and columns are zero indexed.
global row_Key
global col_Key
row_Key = 1
#End of Writing of the Title for individual sheets (KEYS)

#Writing of the Title for individual sheets (RESONA)
worksheet_Resona.write(0, 0, 'Card no')
worksheet_Resona.write(0, 1, 'Name')
worksheet_Resona.write(0, 2, 'Rarity')
worksheet_Resona.write(0, 3, 'Link Suffix')
worksheet_Resona.write(0, 4, 'Effects')
worksheet_Resona.write(0, 5, 'Kana')
worksheet_Resona.write(0, 6, 'Romaji')
worksheet_Resona.write(0, 7, 'Color')
worksheet_Resona.write(0, 8, 'Card Type')
worksheet_Resona.write(0, 9, 'Level')
worksheet_Resona.write(0, 10, 'Power')
worksheet_Resona.write(0, 11, 'Limiting Condition')
worksheet_Resona.write(0, 12, 'Class')
# Start from the first cell. Rows and columns are zero indexed.
global row_Resona
global col_Resona
row_Resona = 1
#End of Writing of the Title for individual sheets (RESONA)

#Writing of the Title for individual sheets (SPELLS)
worksheet_Spell.write(0, 0, 'Card no')
worksheet_Spell.write(0, 1, 'Name')
worksheet_Spell.write(0, 2, 'Rarity')
worksheet_Spell.write(0, 3, 'Link Suffix')
worksheet_Spell.write(0, 4, 'Effects')
worksheet_Spell.write(0, 5, 'Kana')
worksheet_Spell.write(0, 6, 'Romaji')
worksheet_Spell.write(0, 7, 'Color')
worksheet_Spell.write(0, 8, 'Card Type')
worksheet_Spell.write(0, 9, 'Cost')
worksheet_Spell.write(0, 10, 'Limiting Conditon')
# Start from the first cell. Rows and columns are zero indexed.
global row_Spell
global col_Spell
row_Spell = 1
#End of Writing of the Title for individual sheets (SPELLS)

BASE_WIXOSS_URL = 'http://selector-wixoss.wikia.com/'
WIXOSS_WIKI_SUFFIX = 'wiki/'

class WixossSet(Enum):
    #WX01 = 'WX-01_Served_Selector'
    #WX04 = 'WX-04_Infected_Selector'
    #WX22 = 'WX-22_Unlocked_Selector'
    WX21 = 'WX-21_Betrayed_Selector'

class WixossType(Enum):
    LRIG = 'LRIG'
    ARTS = 'ARTS'
    SIGNI = 'SIGNI'
    Spell = 'Spell'
    Key = 'Key'
    Resona = 'Resona'

def get_cards_in_set(wx_set):
    # wx_set WixossSet Enum
    get_card_list_in_set(wx_set)

def get_card_list_in_set(wx_set):
    # wx_set WixossSet Enum
    wx_set_url = BASE_WIXOSS_URL + WIXOSS_WIKI_SUFFIX + wx_set.value
    response = simple_get(wx_set_url)
    html = BeautifulSoup(response, 'html.parser')
    card_list_h2 = html.find(id='Card_List')
    card_list_table = card_list_h2.find_next('table')
    card_list_rows = card_list_table.find_all('tr')
    card_list_rows.pop(0) # First is always just the categories

    for card_list_row in card_list_rows:
        get_card_summary(card_list_row)

def get_card_summary(card_list_row):
    # card_list_row Single row from the card_list_table
    card_summary_tds = card_list_row.find_all('td')
    card_no = card_summary_tds[0].text.strip()
    card_name = card_summary_tds[1].find('a')['title']
    card_link_suffix = card_summary_tds[1].find('a')['href']
    card_rarity = card_summary_tds[2].text.strip()
    card_color = card_summary_tds[3].text.strip()
    global card_type #I want to make this global
    card_type = card_summary_tds[4].text.strip()

    #Excel File will be written here
    global Eng
    Eng = 0
    global row_Lrig
    global row_Signi
    global row_Arts
    global row_Key
    global row_Resona
    global row_Spell
    if card_type == 'LRIG':
        worksheet_Lrig.write(row_Lrig, 0, card_no )
        worksheet_Lrig.write(row_Lrig, 1, card_name )
        worksheet_Lrig.write(row_Lrig, 2, card_rarity )
        worksheet_Lrig.write(row_Lrig, 3, card_link_suffix)
        row_Lrig += 1
    if card_type == 'ARTS':
        worksheet_Arts.write(row_Arts, 0, card_no )
        worksheet_Arts.write(row_Arts, 1, card_name )
        worksheet_Arts.write(row_Lrig, 2, card_rarity )
        worksheet_Arts.write(row_Arts, 3, card_link_suffix)
        row_Arts += 1
    if card_type == 'Key':
        worksheet_Key.write(row_Key, 0, card_no )
        worksheet_Key.write(row_Key, 1, card_name )
        worksheet_Key.write(row_Key, 2, card_rarity )
        worksheet_Key.write(row_Key, 3, card_link_suffix)
        row_Key += 1
    if card_type == 'SIGNI':
        worksheet_Signi.write(row_Signi, 0, card_no )
        worksheet_Signi.write(row_Signi, 1, card_name )
        worksheet_Signi.write(row_Signi, 2, card_rarity )
        worksheet_Signi.write(row_Signi, 3, card_link_suffix)
        row_Signi += 1
    if card_type == 'Spell':
        worksheet_Spell.write(row_Spell, 0, card_no )
        worksheet_Spell.write(row_Spell, 1, card_name )
        worksheet_Spell.write(row_Spell, 2, card_rarity )
        worksheet_Spell.write(row_Spell, 3, card_link_suffix)
        row_Spell += 1
    if card_type == 'Resona':
        worksheet_Resona.write(row_Resona, 0, card_no )
        worksheet_Resona.write(row_Resona, 1, card_name )
        worksheet_Resona.write(row_Resona, 2, card_rarity )
        worksheet_Resona.write(row_Resona, 3, card_link_suffix)
        row_Resona += 1
    #End of writing of excel File
        
    print('')
    print('Card No: ' + card_no )
    print('Name: ' + card_name )
    print('Rarity:' + card_rarity )
    #print('Colour: ' + card_color )
    #print('Type: ' + card_type )
    print('Link Suffix: ' + card_link_suffix )


    card_definer(card_type, card_link_suffix)

def card_definer(type_str, link_suffix):
    get_card_details(link_suffix)

def get_card_details(link_suffix):
    card_url = BASE_WIXOSS_URL + link_suffix
    response = simple_get(card_url)
    html = BeautifulSoup(response, 'html.parser')
    card_cftable = html.find(id='cftable')
    card_container = card_cftable.find(id='container')
    card_info_container = card_cftable.find(id='info_container')
    card_abilities_container = card_info_container.find_next('div', {'class': 'info-extra'})
    card_abilities_langs = card_abilities_container.find_all('table')
    
    if card_abilities_langs:
        for card_abilities_lang in card_abilities_langs:
            card_abilities = card_abilities_lang.find_all('tr')[1].find('td')
            get_card_abilities(card_abilities)

    card_main_info_container = card_info_container.find_next('div', {'class': 'info-main'})
    card_main_info_rows = card_main_info_container.find_all('tr')
    print(card_main_info_rows)

    #Addition to Excel
    col_Lrig=5
    col_Arts=5
    col_Key=5
    col_Spell=5
    col_Signi=5
    col_Resona=5

    for card_main_info_row in card_main_info_rows:
        row_tds = card_main_info_row.find_all('td')
        stat_name = row_tds[0].text.strip()
        stat_value = row_tds[1].text.strip()
        #if stat_name == 'Grow Cost':
         #   content = ''
          #  for ###This is a very big problem. The grow cost colour is missing
           # stat_value=content
        #Upload this to Excel
        if card_type == 'LRIG':
            worksheet_Lrig.write(row_Lrig-1, col_Lrig, stat_value)
            col_Lrig += 1
        if card_type == 'ARTS':
            worksheet_Arts.write(row_Arts-1, col_Arts, stat_value)
            col_Arts += 1
        if card_type == 'Key':
            worksheet_Key.write(row_Key-1, col_Key, stat_value)
            col_Key += 1
        if card_type == 'SIGNI':
            worksheet_Signi.write(row_Signi-1, col_Signi, stat_value)
            col_Signi += 1
        if card_type == 'Spell':
            worksheet_Spell.write(row_Spell-1, col_Spell, stat_value)
            col_Spell += 1
        if card_type == 'Resona':
            worksheet_Resona.write(row_Resona-1, col_Resona, stat_value)
            col_Resona += 1
        #Uploaded to Excel
        print(stat_name + ': ' + stat_value)

    col_Lrig=5
    col_Arts=5
    col_Key=5
    col_Spells=5
    col_Signi=5
    Col_Resona=5
    #Addition to Excel ends


def get_card_abilities(card_abilities):
    global Eng
    if Eng != 0:
        Eng += 1
        return
    print('Effects:')
    content = ''
    for tag in card_abilities:
        if tag.name == 'span':
            content += '[' + tag.text + '] '
        elif tag.name == 'a' and tag.find('img'):
            content += '{' + tag.find('img')['alt'] + '} '
        else:
            if tag == ' ':
                pass
            elif tag == '.\n':
                content = content[:-1] + tag
            elif tag.name:
                content += tag.text + ' ' 
            else:
                content += tag.strip() + ' '
    #Upload this to Excel
    if card_type == 'LRIG':
        worksheet_Lrig.write(row_Lrig-1, 4, content)
        Eng += 1
    if card_type == 'ARTS':
        worksheet_Arts.write(row_Arts-1, 4, content)
        Eng += 1
    if card_type == 'Key':
        worksheet_Arts.write(row_Arts-1, 4, content)
        Eng += 1
    if card_type == 'SIGNI':
        worksheet_Signi.write(row_Signi-1, 4, content)
        Eng += 1
    if card_type ==  'Spell':
        worksheet_Spell.write(row_Spell-1, 4, content)
        Eng += 1
    if card_type == 'Resona':
        worksheet_Resona.write(row_Resona-1, 4, content)
        Eng += 1
    #Uploaded to Excel
    print(content)
    

def simple_get(url):
    try:
        with closing(get(url, stream=True)) as resp:
            if is_good_response(resp):
                return resp.content
            else:
                return None

    except RequestException as e:
        log_error('Error during requests to {0} : {1}'.format(url, str(e)))
        return None


def is_good_response(resp):
    content_type = resp.headers['Content-Type'].lower()
    return (resp.status_code == 200 
            and content_type is not None 
            and content_type.find('html') > -1)


def log_error(e):
    print(e)

def main():
    for wx_set in WixossSet:
        wx_set_cards = get_cards_in_set(wx_set)

if __name__ == "__main__":
    main()

workbook.close()


