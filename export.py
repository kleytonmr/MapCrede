import os
from os.path import isfile, join, basename
import zipfile
import tarfile
import shutil
import patoolib

class Crede:
  def __init__(self):
    self.count = 1
    self.file_old = []

  # Cria diretórios padrão
  def create_folder(self, dir):
    os.mkdir(dir + '/arquivos_excel')
    os.mkdir(dir + '/sem_alteracao')
    os.mkdir(dir + '/extração_manual')

  # Move os arquivos compactados
  def move(self,path_old, path_new1, path_new2):
    for name in os.listdir(path_old):
      if name.endswith("tar.gz"):
        shutil.move(path_old + name, path_new1)
      elif name.endswith("tar"):
        shutil.move(path_old + name, path_new1)
      elif name.endswith("rar"):
        shutil.move(path_old + name, path_new1)
      elif name.endswith("zip"):
        shutil.move(path_old + name, path_new1)
      elif name.endswith("bz2"):
        shutil.move(path_old + name, path_new1)
      elif name.endswith("ods"):
        shutil.move(path_old + name, path_new2)
      elif name.endswith("xls"):
        shutil.move(path_old + name, path_new2)
      elif name.endswith("xlsx"):
        shutil.move(path_old + name, path_new2)
      elif name.endswith("xlsm"):
        shutil.move(path_old + name, path_new2)
      elif name.endswith("xltm"):
        shutil.move(path_old + name, path_new2)
      elif name.endswith("xltx"):
        shutil.move(path_old + name, path_new2)

  # Lista todos os diretórios
  def read_dir(self, dir):
    for nome in os.listdir(dir + '/sem_alteracao'):
      self.tar_xzf(dir,nome)

  # Extrai os arquivos compactados
  def tar_xzf(self, dir, fname):
    try:
      if (fname.endswith("tar.gz")):
          tar = tarfile.open(dir + '/sem_alteracao/' + fname, "r:gz")
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("tar")):
          tar = tarfile.open(dir + '/sem_alteracao/' + fname, "r:")
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("zip")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("rar")):
          patoolib.extract_archive(dir + '/sem_alteracao/' + fname, outdir= dir + '/arquivos_excel')
      elif (fname.endswith("7z")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("ace")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("arj")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("bz2")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("cab")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("gz")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("lz")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("lhz")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("uue")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("xz")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("z")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("zipx")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
      elif (fname.endswith("001")):
          tar = zipfile.ZipFile(dir + '/sem_alteracao/' + fname)
          tar.extractall(dir + '/arquivos_excel')
          tar.close()
    except:
      print('Deu merda: '+ fname)
      self.is_error(dir, fname)

  # Verifica extensão do arquivo
  def verify_ex(self, ex):
    if ex == 'xls' or ex == 'xlsx' or ex == 'xlsm' or ex == 'xltm' or ex == 'xltx' or ex == 'ods' or ex == '.ods':
      return True
    else:
      return False

  # Renomeia os arquivos em excel
  def rename(self, dir):
    folder = '/arquivos_excel/'
    for name in os.listdir(dir + folder):
      n = name.split('.')
      try:
        if self.verify_ex(n[1]):
          os.rename(dir + folder + name, dir + folder + 'Escola_' + str(self.count) + '.' + n[1])
          self.count = self.count + 1
        elif self.verify_ex(n[2]):
          os.rename(dir + folder + name, dir + folder + 'Escola_' + str(self.count) + '.' + n[2])
          self.count = self.count + 1
        elif self.verify_ex(n[3]):
          os.rename(dir + folder + name, dir + folder + 'Escola_' + str(self.count) + '.' + n[3])
          self.count = self.count + 1
      except :
        print(n)

  # Move para extração manual
  def is_error(self, dir, name):
    shutil.move(dir + '/sem_alteracao/' + name, dir + '/extração_manual/')

  # Percorre todos os diretórios e executa todas as funções.
  def map(self):
    for folder in os.listdir('.'):
      print("==>" + folder)
      try:
        self.create_folder(folder + '/')
        self.move(folder + '/', folder + '/sem_alteracao', folder + '/arquivos_excel')
        self.read_dir(folder)
        self.rename(folder)
        self.count = 1
      except:
        pass





