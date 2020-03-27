import win32com.client
import win32api
from pywintypes import com_error
from pathlib import Path

### pip install pypiwin32 if module not found

REPLACE_TXTS = {
	'[organisation]': 'TOP ORGANISATION',
}

excel_path = str(Path(__file__).absolute().parent / 'contract.xlsx')
save_path = str(Path.cwd() / 'contract.pdf')

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
	wb = excel.Workbooks.Add(excel_path)
	sheet = wb.WorkSheets(1)
	sheet.Activate()

	rg = sheet.Range(sheet.usedRange.Address)
	for search_txt, replace_Txt in REPLACE_TXTS.items():
			rg.Replace(search_txt, replace_Txt)

	wb.SaveAs(save_path, FileFormat=57) # 57 == .pdf

except com_error as e:
	print('Failure.', e)
else:
	print('Success.')
finally:
	wb.Close(False)
	excel.Quit()