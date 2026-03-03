import win32print

PRINTER_NAME = "GP-C80250 Series"

handle = win32print.OpenPrinter(PRINTER_NAME)

data = b'\x1b\x40'
data += b'FUNCIONANDO OK IFSUL CAMPUS SAPUCAIA\n\n\n'
data += b'\x1d\x56\x00'

win32print.StartDocPrinter(handle, 1, ("Test", None, "RAW"))
win32print.StartPagePrinter(handle)

win32print.WritePrinter(handle, data)

win32print.EndPagePrinter(handle)
win32print.EndDocPrinter(handle)
win32print.ClosePrinter(handle)

print("Enviado")
