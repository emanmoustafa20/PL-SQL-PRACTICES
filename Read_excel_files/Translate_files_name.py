"""
--------------------------------------------------------------
Class TranslateFilesName is implemented in order to translate
excel files name from Arabic language to English language, 
Also update the database with the translation info
-------------------------------------------------------------
"""
import os
from re import X
import cx_Oracle
import arabic_reshaper
from bidi.algorithm import get_display

class TranslateFilesName:
   
    # init method or constructor 
    def __init__(self,root_directory):
        self.root_directory = root_directory
        self.excel_directories_path = []
        self.excel_directories = []
        self.excel_files = []
   
    # get all excel files and their directories
    
    def get_files_and_dirs(self):
      for (root,dirs,files) in os.walk(self.root_directory):     
        for j in files:     
          self.excel_directories_path.append(root)
          self.excel_files.append(j)
        for j in dirs:
          self.excel_directories.append(dirs)
         
      #print(len(self.excel_directories_path))
     
    


   
    def translate_file_names(self):
        for i in range(0,len(self.excel_files)):
            old_file_path =  os.path.join( self.excel_directories_path[i],self.excel_files[i])
            print(old_file_path)
            new_file_name = 'F' + str(i) +'.xlsx'
            new_file_path = os.path.join(self.excel_directories_path[i],new_file_name)
            print(new_file_path)
            os.rename(old_file_path,new_file_path)
            self.insert_translated_file_names(i,self.excel_files[i],new_file_name,old_file_path,self.excel_directories_path[i])
            os.rename(old_file_path,new_file_path)
            
   
     
    def translate_dir_names(self):
          for i in range(0,len(self.excel_directories[1])):
             new_dir_name = 'D' + str(i)
             orig_dir_name = self.excel_directories[1][i]
             self.insert_translated_dir_names(orig_dir_name, new_dir_name)
            #dirs returns two dim list, due to having one root file(original) and one subfile claims_sheets, it's 1 and i  
             old_file_path = os.path.join(self.root_directory,self.excel_directories[1][i])
             new_file_path = os.path.join(self.root_directory,new_dir_name)
             
             
             #os.rename(old_file_path,new_file_path)
             reshaped_text = arabic_reshaper.reshape(self.excel_directories[1][i])    # correct its shape
             reshaped_text_displayable = get_display(reshaped_text) 
             #print(reshaped_text_displayable)

    def insert_translated_file_names(self,id,orig_name,trans_name,file_path,dir_path):
        dsn_tns = cx_Oracle.makedsn('Eman-Mostafaa.uhia.local', 1521, service_name='orclpdb.uhia.local') # if needed, place an 'r' before any parameter in order to address special characters such as '\'.
        conn = cx_Oracle.connect(user=r'CLAIMS', password='claims', dsn=dsn_tns) # if needed, place an 'r' before any parameter in order to address special characters such as '\'. For example, if your user name contains '\', you'll need to place 'r' before the user name: user=r'User Name'
        c = conn.cursor()
        q = ("insert into CLM_DIRS_AND_FILES(FILE_ID,TRANSLATED_FILE_NAME,ORIGINAL_FILE_NAME,DIRECTORY_PATH,FILE_PATH) values (:ID,:TRANS,:ORIG,:DIR_PATH,:PATH)")
        c.execute(q,ID = id, TRANS = trans_name, ORIG = orig_name, PATH = file_path,DIR_PATH = dir_path)
        conn.commit()

    def insert_translated_dir_names(self,orig_file_name,trans_dir_name):
        dsn_tns = cx_Oracle.makedsn('Eman-Mostafaa.uhia.local', 1521, service_name='orclpdb.uhia.local') # if needed, place an 'r' before any parameter in order to address special characters such as '\'.
        conn = cx_Oracle.connect(user=r'CLAIMS', password='claims', dsn=dsn_tns) # if needed, place an 'r' before any parameter in order to address special characters such as '\'. For example, if your user name contains '\', you'll need to place 'r' before the user name: user=r'User Name'
        c = conn.cursor()
        c.execute('select DIRECTORY_ORIGINAL_NAME from CLM_DIRS') 
        #c.execute(x, DIR_NAME =orig_file_name)
        for row in c:
          insert_flag =False
          q = ("UPDATE CLM_DIRS_AND_FILES SET TRANSLATED_DIR_NAME =: TRANS_DIR_NAME WHERE DIRECTORY_NAME =:orig_file_name")
          c.execute(q, TRANS_DIR_NAME =trans_dir_name, orig_file_name =orig_file_name)
          conn.commit()
        if insert_flag:
          q = ("insert into CLM_DIRS(DIRECTORY_ORIGINAL_NAME,DIRECTORY_TRANSLATED_NAME, DIRECTORY_PATH)  VALUES(:DIR_ORIG_NAME, :DIR_TRAN_NAME, :DIR_PATH)")
          c.execute(q, TRANS_DIR_NAME =trans_dir_name, orig_file_name =orig_file_name)
          conn.commit()
    

p = TranslateFilesName('D:\ORIGINAL_IN_ENGLISH')
p.get_files_and_dirs()
p.translate_file_names()
#p.translate_dir_names()
