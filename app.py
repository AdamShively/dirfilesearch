"""
This program is the GUI
used in conjunction with
search.py.
Author: Adam Shively
Version: 1.2
Since: 01-22-2022
"""

import os
import clipboard
import PySimpleGUI as sg
from search import DirSearch

#Creates main window
def create_main_window():

    path_box_layout = ['Delete', 'Copy', 'Paste']
    search_box_layout = ['Delete ', 'Copy' , 'Paste ']

    #Layout for window  
    search_layout = [
        [sg.Text('Path:', size =(15, 1)), sg.InputText('', size=100, key='-PATH_BOX-', right_click_menu=['', path_box_layout])],
        [sg.Text('Search Phrase:', size =(15, 1)), sg.InputText('', size=30, key='-SEARCH_BOX-', right_click_menu=['', search_box_layout]), 
            sg.Radio('Search: Exact Phrase', "RADIO", default=True, key='-EXACT-'), 
            sg.Radio('Search: Any Word in Phrase', "RADIO", default=False, key='-ANY-'), 
            sg.Radio('Search: All Words in Phrase Present', "RADIO", default=False, key='-ALL-')],
        [sg.Button('Search'), sg.Button('Exit'), 
            sg.Checkbox('Search docx Files', default=True, key='-WORD-'), sg.Checkbox('Search txt Files', default=True, key='-TXT-')],
        [sg.Text('_____'  * 30, size=(124, 1))]
        ]

    results_layout = [[sg.Listbox(values=[None], 
                size=(140, 12), key='-RESULTS-', enable_events=True)],sg.Text('_____'  * 30, size=(125, 1))]
    results_buttons = [sg.Button('Open'), sg.Text('Files That Meet Search Parameters:')]

    error_layout = [[sg.Text('An Error Occurred With The Following:')],
                    [sg.Listbox(values=[None], size=(140, 8), key='-ERROR_LIST-', enable_events=True)]]

    return sg.Window('Search Folders and Files', 
            [search_layout, results_buttons,
            results_layout, error_layout],
            size=(1100, 650), finalize=True)

def get_search_type(values):                           
    if values['-ANY-']:
        return 'any' 
            
    elif values['-ALL-']:
        return 'all'  
            
    else:
        return 'exact'  #Default search type

def get_file_type(values):
    if values['-WORD-'] and values['-TXT-']:
        return 'both'
            
    elif values['-WORD-']:
        return 'word'  
            
    elif values['-TXT-']:
        return 'txt'

    else:
        return ''  #Nothing selected

#main function
def main():
    window = create_main_window()   #Create window
    found_list = []

    window['Open'].update(disabled=True)
    
    while True:
        event, values = window.read()

        path_box = values['-PATH_BOX-'].replace('\\', '//')
        if  len(path_box) > 0 and path_box[-1] != '/':
            path_box = f'{path_box}/'

        if event in (sg.WIN_CLOSED, 'Exit'):    #Exit window
                break
       
        #---------------SEARCH BUTTON PRESSED------------------
        if event in ('Search'):

            search_type = get_search_type(values)                          
            file_type = get_file_type(values)

            targets = values['-SEARCH_BOX-']   #Values to search for
            ds = DirSearch(path_box, targets, search_type, file_type)    #Create instance of DirSearch class
            
            if not os.path.isdir(path_box) and not DirSearch.is_word(path_box) and not DirSearch.is_txt(path_box):  
                sg.popup_error('ERROR',"Enter a valid directory or docx file path.", keep_on_top=True)    #Unacceptable path entered
                
            elif not targets:                                   
                sg.popup_error('ERROR',"No search phrase has been entered.", keep_on_top=True)  #No search phrase entered
            
            else:
                is_empty_dir = len(os.listdir(path_box)) == 0   #Check if empty directory
                if os.path.isdir(path_box):   #Directory path was entered
                    
                    if is_empty_dir:  #Check if empty directory
                        sg.popup_error('ERROR',"Entered directory is empty.", keep_on_top=True)
                    
                    else:
                        found_list = ds.get_files()
                
                else:                                                   
                    found_list = [path_box] if ds.text_search() else []  #Word document path was entered                                 

                window['-RESULTS-'].update(found_list[0], set_to_index=[0])
                if found_list[0]:                                         #Search results list non-empty
                    window['Open'].update(disabled=False)
                else:
                    window['Open'].update(disabled=True)

                window['-ERROR_LIST-'].update(found_list[1])
                
                if not is_empty_dir and not found_list[0]:   #No files containing search phrase and directory not empty
                    sg.popup('No Files Found',"No files found that meet search parameters.", keep_on_top=True)

        #---------------RIGHT CLICK MENU------------------
        if event in ('Paste', 'Paste '):
            which_box = window['-PATH_BOX-'] if event == 'Paste' else window['-SEARCH_BOX-']
            try:   #Deletes selected before pasting
                which_box.Widget.delete("sel.first", "sel.last")
                text = clipboard.paste()    
                which_box.Widget.insert("insert", text)   #Paste text
            
            except:                     
                text = clipboard.paste()    
                which_box.Widget.insert("insert", text)   #Paste text
            
        if event in ('Copy', 'Copy '):
            which_box = window['-PATH_BOX-'] if event == 'Copy' else window['-SEARCH_BOX-']
            try:
                text = which_box.Widget.selection_get()
                clipboard.copy(text)  #Add selected to clipboard

            except Exception as e:                   
                sg.popup_error('ERROR', f'{e}\nSelect portion of text to be copied.', keep_on_top=True)

        if event in ('Delete', 'Delete '):  
            which_box = window['-PATH_BOX-'] if event == 'Delete' else window['-SEARCH_BOX-']
            try:
                which_box.Widget.delete("sel.first", "sel.last")  #Delete selected
            
            except Exception as e:                   
                sg.popup_error('ERROR', f'{e} \n Select portion of text to be deleted.', keep_on_top=True)

        #---------------RESULTS FILES------------------
        if event in ('Open',):                                  #Open file selected
            open_file = values['-RESULTS-'][0]
            os.startfile(open_file)

    window.close()

if __name__ == "__main__":
    main()