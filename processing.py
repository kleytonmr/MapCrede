from os.path import isfile, join, basename
import os, json, shutil
import pandas as pd
from lxml import etree

class Processing:
  def __init__(self):
    pass
  
  # Lista todos os diretórios
  def ods(self, dir):
    for current_file in os.listdir(dir):
      if current_file.endswith("xls") or current_file.endswith("xlsx"):
        print(dir + '/' + current_file)
        try:
          file = pd.ExcelFile(dir + '/' + current_file)
          writer = pd.ExcelWriter(dir + '/' + current_file)
          print(current_file)
          for x in file.sheet_names:
            df = pd.read_excel(file, sheet_name=x)
            df['nome_original_' + 'CREDE 13'] = current_file
            df.to_excel(writer, x)
            writer.save()
        except:
          print('Error -------> ' + current_file)
          shutil.move(dir + '/' + current_file, 'CREDE 13/bases não extraidas/' + current_file)
  
  def read_dir(self, dir):
    for current_file in os.listdir(dir):
      if current_file.endswith("xls") or current_file.endswith("xlsx"):
        try:
          file = pd.ExcelFile(dir + '/' + current_file)
          writer = pd.ExcelWriter(dir + '/bases extraidas/' + current_file)
          print(current_file)
          for x in file.sheet_names:
            df = pd.read_excel(file, sheet_name=x)
            df['nome_original_' + dir] = current_file
            df.to_excel(writer, x)
            writer.save()
        except:
          print('Error -------> ' + current_file)

 # Percorre todos os diretórios e executa todas as funções. 
  def map(self):
    # self.ods('CREDE 13/Exrtação ods')
    print('CREDE 13')
    self.read_dir('CREDE 13')
      
      
  
  
    
  
    