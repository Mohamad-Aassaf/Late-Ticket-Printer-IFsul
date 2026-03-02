import win32print

printers = win32print.EnumPrinters(2)

for printer in printers:
    print(printer[2])
