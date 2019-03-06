import os
from os.path import isfile, join, basename
import zipfile
import tarfile
import shutil
import patoolib

class Crede:
  def __init__(self):
    pass

  # Cria diretórios padrão
  def create_folder(self, dir):
    os.mkdir(dir + '/bases extraidas')
    os.mkdir(dir + '/bases não extraidas')

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
    for nome in os.listdir(dir):
      self.tar_xzf(dir,nome)

  # Extrai os arquivos compactados
  def tar_xzf(self, dir, fname):
    try:
      if (fname.endswith("tar.gz")):
          tar = tarfile.open(dir + '/' + fname, "r:gz")
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("tar")):
          tar = tarfile.open(dir + '/' + fname, "r:")
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("zip")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("rar")):
          patoolib.extract_archive(dir + '/' + fname, outdir= dir)
      elif (fname.endswith("7z")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("ace")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("arj")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("bz2")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("cab")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("gz")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("lz")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("lhz")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("uue")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("xz")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("z")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("zipx")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
      elif (fname.endswith("001")):
          tar = zipfile.ZipFile(dir + '/' + fname)
          tar.extractall(dir)
          tar.close()
    except:
      print('Movido P/ bases não extraidas: '+ fname)
      self.is_error(dir, fname)

  # Move para extração manual
  def is_error(self, dir, name):
    shutil.move(dir + '/' + name, dir + '/bases não extraidas/')

  # Percorre todos os diretórios e executa todas as funções.
  def map(self):
    for folder in os.listdir('.'):
      print("==> " + folder)
      try:
        self.create_folder(folder + '/')
        self.read_dir(folder)
      except:
        pass



