import win32com.client
from pywintypes import com_error
from pathlib import Path

REPLACE_TXTS = {
	'[organisation]': 'TOP ORGANISATION',
}

nowdir = Path(__file__).absolute().parent

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False

try:
	wb = excel.Workbooks.Add(str(nowdir / 'contract.xlsx'))
	sheet = wb.WorkSheets(1)
	sheet.Activate()

	rg = sheet.Range(sheet.usedRange.Address)
	for search_txt, replace_Txt in REPLACE_TXTS.items():
			rg.Replace(search_txt, replace_Txt)

	print(rg)
	wb.SaveAs(str(nowdir / 'new.xlsx'))

except com_error as e:
	print('Failure.', e)
else:
	print('Success.')
finally:
	wb.Close(False)
	excel.Quit()