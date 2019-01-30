import os
from os.path import isfile, join, basename
import zipfile
import tarfile
import shutil

class Crede:
  def __init__(self):
    self.count = 1
    self.file_old = []

  # Cria diretórios padrão
  def create_folder(self, dir):
    try:
      os.mkdir(dir + '/arquivos_excel')
      os.mkdir(dir + '/sem_alteracao')
      os.mkdir(dir + '/extração_manual')
    except:
      print('Pastas já existem!')

  # Move os arquivos compactados
  def move(self,path_old, path_new1, path_new2):
    for name in os.listdir(path_old):
      if name.endswith("tar.gz"):
        shutil.move(path_old + name, path_new1)
      elif name.endswith("tar"):
        shutil.move(path_old + name, path_new1)
      elif name.endswith("zip"):
        shutil.move(path_old + name, path_new1)
      elif name.endswith("ods"):
        shutil.move(path_old + name, path_new2)
      elif name.endswith("xls"):
        shutil.move(path_old + name, path_new2)
      elif name.endswith("xlsx"):
        shutil.move(path_old + name, path_new2)
      elif name.endswith("xlsm"):
        shutil.move(path_old + name, path_new2)

  # Lista todos os diretórios
  def read_dir(self, dir):
    for nome in os.listdir(dir):
      self.tar_xzf(dir,nome)

  # Extrai os arquivos compactados
  def tar_xzf(self, dir, fname):
    try:
      if (fname.endswith("tar.gz")):
          tar = tarfile.open(dir + fname, "r:gz")
          tar.extractall('CREDE 13/arquivos_excel')
          tar.close()
      elif (fname.endswith("tar")):
          tar = tarfile.open(dir + fname, "r:")
          tar.extractall('CREDE 13/arquivos_excel')
          tar.close()
      elif (fname.endswith("zip")):
          tar = zipfile.ZipFile(dir + fname)
          tar.extractall('CREDE 13/arquivos_excel')
          tar.close()
    except:
      print('Deu merda: '+ fname)
      self.file_old.append(fname)

  def rename(self, dir):
    for name in os.listdir(dir):
      n = name.split('.')
      try:
        os.rename(dir + name, dir + 'Escola_' + str(self.count) + '.' + n[2])
        self.count = self.count + 1
      except:
        os.rename(dir + name, dir + 'Escola_' + str(self.count) + '.' + n[1])
        self.count = self.count + 1

  def is_error(self, dir):
    for name in self.file_old:
      shutil.move(dir + name, 'CREDE 13/extração_manual')

  def map(self):
    for folder in os.listdir('.'):
      try:
        self.create_folder(folder + '/')
        self.move(folder + '/', folder + '/sem_alteracao', folder + '/arquivos_excel')
        self.read_dir(folder + '/sem_alteracao/')
        self.rename(folder + '/arquivos_excel/')
        self.is_error(folder + '/sem_alteracao/')
      except:
        pass





