import xlwings as xw, psutil

# Nowy plik Excela
wb = xw.Book()

# Import modułu VBA
wb.api.VBProject.VBComponents.Import(r"makro.bas")

# Zapisz jako plik z makrami (.xlsm)
wb.save(r"plik_z_makrem.xlsm")

# Uruchom makro
makro = wb.macro("HelloWorld")
makro()

wb.close()