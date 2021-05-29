"""
--------------------------------------------------------------
Class TranslateFilesName is implemented in order to translate
excel files name from Arabic language to English language, 
Also update the database with the translation info
-------------------------------------------------------------
"""
import os
import cx_Oracle

class TranslateFilesName:
   
    # init method or constructor 
    def __init__(self,root_directory):
        self.root_directory = root_directory
        self.excel_directories = []
        self.excel_files = []
   
    # get all excel files and their directories
    
    def get_files_and_dirs(self):
      for (root,dirs,files) in os.walk(self.root_directory):     
        for j in files:     
          self.excel_directories.append(root)
          self.excel_files.append(j)
      print(len(self.excel_directories))
      print(len(self.excel_files))
    

    def translate_file_names(self):
        for i in range(0,len(self.excel_files)):
            old_file_path =  os.path.join( self.excel_directories[i],self.excel_files[i])
            print(old_file_path)
            new_file_name = 'F' + str(i) +'.xlsx'
            new_file_path = os.path.join(self.excel_directories[i],new_file_name)
            print(new_file_path)
            #os.rename(old_file_path,new_file_path)
            self.connect_to_oracle(i,self.excel_files[i],new_file_name,old_file_path,self.excel_directories[i])
            os.rename(old_file_path,new_file_path)
            


    def connect_to_oracle(self,id,orig_name,trans_name,file_path,dir_name):
        dsn_tns = cx_Oracle.makedsn('Eman-Mostafaa.uhia.local', 1521, service_name='XE') # if needed, place an 'r' before any parameter in order to address special characters such as '\'.
        conn = cx_Oracle.connect(user=r'CLAIMS', password='claims', dsn=dsn_tns) # if needed, place an 'r' before any parameter in order to address special characters such as '\'. For example, if your user name contains '\', you'll need to place 'r' before the user name: user=r'User Name'
        c = conn.cursor()
        q = ("insert into CLM_DIRS_AND_FILES(FILE_ID,TRANSLATED_FILE_NAME,ORIGINAL_FILE_NAME,DIRECTORY_NAME,FILE_PATH) values (:ID,:TRANS,:ORIG,:DIR,:PATH)")
        c.execute(q,ID = id, TRANS = trans_name, ORIG = orig_name, PATH = file_path,DIR = dir_name)
        conn.commit()


    

p = TranslateFilesName('D:\orig')
p.get_files_and_dirs()
p.translate_file_names()