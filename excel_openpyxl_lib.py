from openpyxl import load_workbook, Workbook

class excel_file():

	def __init__(self, filepath, next_line=1):
		"""
			filepath est le chemin complet du fichier
			next_line est la ligne suivante à laquelle écrire, par défaut, la première
		"""
		self.filepath = filepath
		self.next_line = next_line


	def create_workbook(self):
		wb = Workbook()
		wb.save(self.filepath)

	def open(self, sheet='first'):
		self.wb = load_workbook(self.filepath)
		if sheet == 'first':
			self.ws = self.wb.active
		else:
			raise TypeError("Valeur de 'sheet' incorrecte")

	def write(self, data, from_cell=1, from_line=None):
		""" Ecriture dans le fichier, 
			Données à écrire:
				- Si données n'est pas une liste ou un tuple, on écrit une seule case
				- Si données est une liste ou un tuple à un seul niveau, on écrit une ligne
				- Si données est une liste ou un tuple à deux niveaux, on écrit plusieurs lignes
			Ligne à écrire (une ligne libre est une ligne dont la première cellule est vide):
				- si from_line est null on écrit à la ligne suivante enregistrée dans self.next_line
				- si line est à first_free, on cherche la première ligne libre
				- si line est à next_free, on cherche la ligne libre suivante après self.next_line
		"""
		
		# On défini la ligne à écrire
		if from_line is None:	
			from_line = self.next_line
		elif from_line == "first_free":
			while self.ws.cell(1, 1).value is not None:
				from_line += 1
		elif from_line == 'next_free':
			while self.ws.cell(self.next_line, 1).value is not None:
				from_line += 1
		elif from_line.isdigit():
			pass
		else:
			raise TypeError("Valeur de ligne à écrire incorrecte")

		# On défini la taille des données à écrire
		# data_size est une liste qui contiendra les deux dimentions des données à écrite [ligne, colonne]
		if type(data) not in (tuple, list):
			# Cellule unique
			data_size = [1,1]
		elif type(data[0]) not in (tuple, list):
			# On défini si c'est une liste à deux dimentions en regardant le premier élément de la liste
			# Liste à une dimention
			data_size = [1, len(data)]
		else:
			# Liste à deux dimentions
			data_size = [len(data[0]), len(data)]


		# On écrit les données en fonction des tailles définies 
		raise SyntaxError("Programmation à compléter")


	def append(self, data):
		""" Ajout d'une ou plusieurs lignes directement à la suite """
		self.ws.append(data)


		




		

	def auto_size(self):
		""" Redimentionnement des colonnes à la taille du contenu """
		for column_cells in self.ws.columns:
		    length = max(len(str(cell.value)) for cell in column_cells)
		    self.ws.column_dimensions[column_cells[0].column_letter].width = length


	def close(self):
		self.wb.save(self.filepath)
		self.wb.close()
