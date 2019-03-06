from os.path import isfile, join, basename
import os, json, shutil
import pandas as pd
import xlwt, ezodf
from lxml import etree

class Processing:
  def __init__(self):
    pass
  
  # Lista todos os diretórios
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
          shutil.move(dir + '/' + current_file, dir + '/bases não extraidas/' + current_file)

  def list_dir(self):
    list_files = []
    writer = pd.ExcelWriter('out.xlsx')
    
    for current_file in os.listdir('CREDE 30/bases extraidas'):
      if current_file.endswith("xls") or current_file.endswith("xlsx"):
        list_files.append(current_file)
      
    df = pd.DataFrame(data=list_files)
    df.to_excel(writer,'CREDE 30')
    writer.save()

 # Percorre todos os diretórios e executa todas as funções. 
  def map(self):
    for folder in os.listdir('.'):
      try:
        print(folder)
        self.read_dir(folder)
      except:
        print('Non-numeric data found in the file.')
  
  
    
  
    