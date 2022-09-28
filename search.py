"""
This program is the searcher
used in conjunction with
GUI.py.
Author: Adam Shively
Version: 1.2
Since: 01-22-2022
"""

import os
import docx2txt
import PySimpleGUI as sg

class DirSearch:
    
    #Initialization function
    def __init__(self, directory_path, targets, search_type, file_type):
        self.directory_path = directory_path
        self.targets = targets
        self.search_type = search_type
        self.file_type = file_type

    #Extracts word/txt documents from directories
    #Adds to appropriate lists
    #Returns lists
    def get_files(self):
        stack = [self.directory_path]
        file_collection = []
        unassessable_collection = []

        while stack:
            current = stack.pop()
            try:
                dir_list = os.listdir(current)
            except Exception as e:
                unassessable_collection.append(e)           #Directory could not be searched

            for item in dir_list:
                
                folder_or_file_path = f'{current}{item}'            #Create path name for item in directory
                is_directory = os.path.isdir(folder_or_file_path)
                is_valid_file = DirSearch.is_word(folder_or_file_path) or DirSearch.is_txt(folder_or_file_path)
                
                if is_valid_file:                      #If word doc or txt file

                    try:
                        search = self.text_search(folder_or_file_path)
                        if search:
                            file_collection.append(folder_or_file_path)    #Search successful

                    except Exception as e:

                        unassessable_collection.append(f'{e}: {folder_or_file_path}')     #File could not be searched
                
                elif is_directory:
                    stack.append(f'{folder_or_file_path}/')    #Add to stack if directory
        
        return file_collection, unassessable_collection

    #Check if file is a docx file and not empty
    def is_word(path_name):

        _ , file_extension = os.path.splitext(path_name)
        return file_extension == '.docx' #and 
            #os.stat(path_name).st_size > 0)

    #Check if file is a txt file
    def is_txt(path_name):

        _ , file_extension = os.path.splitext(path_name)
        return file_extension == '.txt'

    #Search for entered search phrase
    def text_search(self, path):
        text = ''

        if self.file_type == 'word' and DirSearch.is_word(path):  #Type chosen is docx and file is docx
            text = docx2txt.process(path).lower()


        elif self.file_type == 'txt' and DirSearch.is_txt(path):    #Type chosen is txt and file is txt
            
            with open(path) as f:
                text = f.read()

        elif self.file_type == 'both':                          #Type chosen is docx and txt
            
            if DirSearch.is_word(path):                         #Check if file is docx
                text = docx2txt.process(path).lower()
                
            elif DirSearch.is_txt(path):                        #Check if file is txt
                with open(path) as f:
                    text = f.read()

        #Base search on chosen search type
        if self.search_type == 'any':
            return self.any_part_search(text)
        
        elif self.search_type == 'all':
            return self.all_present_search(text)
        
        else:
            return self.exact_phrase_search(text)  #Default choice

    #Searches for files containing exact entered phrase
    def exact_phrase_search(self, text):
        target = self.targets.lower()        
        
        if target not in text:                          #Check that word in search phrase is present
            return False
        return True
    
    #Searches for files containing any words in entered phrase (at least one)
    def any_part_search(self, text):
        targets = self.targets.lower().split()
        
        for target in targets:                       
            if target in text:
                return True                      #Check that any part of search string is present
        
        return False

    #Searches for files containing all words in entered phrase (do not have to be in order)
    def all_present_search(self, text):
        targets = self.targets.lower().split()               #Attempt to get text from file
        
        for target in targets:                     
            if target not in text:                           #Check that word in search string is present
                return False
        
        return True