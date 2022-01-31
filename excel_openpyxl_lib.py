from openpyxl import load_workbook, Workbook

class excel_file():

	def __init__(self, filepath):
		self.filepath = filepath


	def create_workbook(self):
		wb = Workbook()
		wb.save(self.filepath)

	def open(self, sheet='first'):
		self.wb = load_workbook(self.filepath)
		if sheet == 'first':
			self.ws = self.wb.active
		else:
			raise TypeError("Valeur de 'sheet' incorrecte")

	def write(self, data, from_cell=1, line=None):
		""" Ecriture dans le fichier, si line est null on écrit à la ligne vide suivante """
		if line is None:	
			line = 1
			while self.ws.cell(line, 1).value is not None:
				line += 1
		for cell, item in enumerate(data):
			self.ws.cell(line, cell + from_cell).value = item

	def auto_size(self):
		""" Redimentionnement des colonnes à la taille du contenu """
		for column_cells in self.ws.columns:
		    length = max(len(str(cell.value)) for cell in column_cells)
		    self.ws.column_dimensions[column_cells[0].column_letter].width = length


	def close(self):
		self.wb.save(self.filepath)
		self.wb.close()
