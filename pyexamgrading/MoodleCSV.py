#	pyexamgrading - Manage grade computation of university exams
#	Copyright (C) 2024-2024 Johannes Bauer
#
#	This file is part of pyexamgrading.
#
#	pyexamgrading is free software; you can redistribute it and/or modify
#	it under the terms of the GNU General Public License as published by
#	the Free Software Foundation; this program is ONLY licensed under
#	version 3 of the License, later versions are explicitly excluded.
#
#	pyexamgrading is distributed in the hope that it will be useful,
#	but WITHOUT ANY WARRANTY; without even the implied warranty of
#	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#	GNU General Public License for more details.
#
#	You should have received a copy of the GNU General Public License
#	along with pyexamgrading; if not, write to the Free Software
#	Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
#
#	Johannes Bauer <JohannesBauer@gmx.de>

import csv
import enum

class MoodleCSVColumn(enum.IntEnum):
	ID = enum.auto()
	FirstName = enum.auto()
	LastName = enum.auto()
	CompleteName = enum.auto()
	Email = enum.auto()
	Organization = enum.auto()
	OrganizationalUnit = enum.auto()
	Status = enum.auto()
	Score = enum.auto()
	MaxScore = enum.auto()
	MayChangeScore = enum.auto()
	LastChangeInput = enum.auto()
	LastChangeGrade = enum.auto()
	Feedback = enum.auto()

class CSVRow():
	def __init__(self, moodle_csv: "MoodleCSV", row_index: int):
		self._moodle_csv = moodle_csv
		self._row_index = row_index

	def update(self, field_dict: dict[MoodleCSVColumn | str, str]):
		for (field, value) in field_dict.items():
			self[field] = value

	@property
	def known_fields(self):
		return { field: self._moodle_csv._raw_get(self._row_index, col_index) for (field, col_index) in self._moodle_csv._known_field_col_indices.items() }

	@property
	def unknown_fields(self):
		return { field: self._moodle_csv._raw_get(self._row_index, col_index) for (field, col_index) in self._moodle_csv._unknown_field_col_indices.items() }

	def to_dict(self):
		fields = self.unknown_fields
		fields.update(self.known_fields)
		return fields

	def __getitem__(self, field: MoodleCSVColumn | str):
		col_index = self._moodle_csv._get_column_index(field)
		return self._moodle_csv._raw_get(self._row_index, col_index)

	def __setitem__(self, field: MoodleCSVColumn | str, value: str):
		col_index = self._moodle_csv._get_column_index(field)
		return self._moodle_csv._raw_set(self._row_index, col_index, value)

	def __str__(self):
		return str(self.to_dict())

class MoodleCSV():
	_COLUMN_NAMES = {
		"ID":								MoodleCSVColumn.ID,
		"Vorname":							MoodleCSVColumn.FirstName,
		"Nachname":							MoodleCSVColumn.LastName,
		"Institution":						MoodleCSVColumn.Organization,
		"Abteilung":						MoodleCSVColumn.OrganizationalUnit,
		"Vollständiger Name":				MoodleCSVColumn.CompleteName,
		"E-Mail-Adresse":					MoodleCSVColumn.Email,
		"Status":							MoodleCSVColumn.Status,
		"Bewertung":						MoodleCSVColumn.Score,
		"Bestwertung":						MoodleCSVColumn.MaxScore,
		"Bewertung kann geändert werden":	MoodleCSVColumn.MayChangeScore,
		"Zuletzt geändert (Abgabe)":		MoodleCSVColumn.LastChangeInput,
		"Zuletzt geändert (Bewertung)":		MoodleCSVColumn.LastChangeGrade,
		"Feedback als Kommentar":			MoodleCSVColumn.Feedback,
	}

	def __init__(self, filename: str):
		with open(filename, encoding = "utf-8-sig") as f:
			self._data = list(csv.reader(f))

		header = self._data[0]
		self._known_field_col_indices = { }
		self._unknown_field_col_indices = { }
		self._index_to_field_name = { }

		for (col_index, col_text) in enumerate(header):
			if col_text in self._COLUMN_NAMES:
				key = self._COLUMN_NAMES[col_text]
				self._known_field_col_indices[key] = col_index
			else:
				key = col_text
				self._unknown_field_col_indices[key] = col_index
			self._index_to_field_name[col_index] = key

	@property
	def column_count(self):
		return len(self._data[0])

	def _get_column_index(self, column: MoodleCSVColumn | str):
		if isinstance(column, MoodleCSVColumn):
			return self._known_field_col_indices[column]
		else:
			return self._unknown_field_col_indices[column]

	def _raw_get(self, row_index: int, col_index: int):
		return self._data[row_index][col_index]

	def _raw_set(self, row_index: int, col_index: int, value: str):
		self._data[row_index][col_index] = str(value)

	def have_column(self, column: MoodleCSVColumn | str):
		return (column in self._known_field_col_indices) or (column in self._unknown_field_col_indices)

	def write_csv(self, filename: str):
		with open(filename, "w") as f:
			writer = csv.writer(f)
			for row in self._data:
				writer.writerow(row)

	def __iter__(self):
		for row_index in range(1, len(self._data)):
			yield CSVRow(self, row_index)

if __name__ == "__main__":
	mcsv = MoodleCSV("TINF23CS_Kryptologie Bewertungen-20240612_1700-comma_separated.csv")
	for entry in mcsv:
		print(entry.unknown_fields)
