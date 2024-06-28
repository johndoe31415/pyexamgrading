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

import odsexport
from .GradingScheme import GradingSchemeType

class ODSExporter():
	def __init__(self, exam: "Exam", entries: list, stats):
		self._exam = exam
		self._entries = entries
		self._stats = stats
		self._doc = None
		self._sheets = { }
		self._styles = {
			"heading": odsexport.CellStyle(font = odsexport.Font(bold = True)),
			"failed": odsexport.CellStyle(background_color = "#e74c3c"),
			"barely_failed": odsexport.CellStyle(background_color = "#9b59b6"),
			"exceptional": odsexport.CellStyle(background_color = "#2ecc71"),
			"#.#":  odsexport.CellStyle(data_style = odsexport.DataStyle.fixed_decimal_places(1)),
			"#.##":  odsexport.CellStyle(data_style = odsexport.DataStyle.fixed_decimal_places(2)),
			"#":  odsexport.CellStyle(data_style = odsexport.DataStyle.fixed_decimal_places(0)),
			"#.#%":  odsexport.CellStyle(data_style = odsexport.DataStyle.percent_fixed_decimal_places(1)),
			"#%":  odsexport.CellStyle(data_style = odsexport.DataStyle.percent_fixed_decimal_places(0)),
		}
		self._cells = {	}

	@property
	def sheet_results(self):
		return self._sheets["results"]

	@property
	def sheet_grade_overview(self):
		return self._sheets["grade_overview"]

	@property
	def style_heading(self):
		return self._styles["heading"]

	def _percent_to_grade_formula(self, percent_cell: "Cell"):
		if self._exam.grading_scheme.grading_scheme_type in [ GradingSchemeType.GermanUniversityLinear, GradingSchemeType.GermanUniversityCutoff ]:
			# Cutoff ends up exactly on 4.05
			formula = f"1 + 3.05*({self._cells['cutoff_high']:a}-{percent_cell})/({self._cells['cutoff_high']:a}-{self._cells['cutoff_low']:a})"

			# Clamp
			formula = odsexport.Formula.clamp(formula, "1.0", "5.0")

			# Round to even with 1 decimal point so 4.05 becomes 4.1, but
			# 4.0500001 becomes 4.1
			formula = odsexport.Formula.round_half_to_even(formula, digits = 1)
		else:
			raise NotImplementedError(self._exam.grading_scheme.grading_scheme_type)
		return odsexport.Formula(formula)

	def _points_to_percent_formula(self, points_sum_cell: "Cell"):
		return odsexport.Formula(f"{points_sum_cell}/{self._exam.structure.max_points:f}")

	def _populate_grade_overview(self):
		sheet = self.sheet_grade_overview
		writer = sheet.writer()

		for col_no in range(3):
			sheet.style_column(col_no, odsexport.ColStyle(width = "3cm"))

		if self._exam.grading_scheme.grading_scheme_type in [ GradingSchemeType.GermanUniversityLinear, GradingSchemeType.GermanUniversityCutoff ]:
			writer.write("Bestehensgrenze (%):", style = self.style_heading)
			writer.skip().write(float(self._exam.grading_scheme.parameters["cutoff_low"]), style = self._styles["#%"]).advance()
			self._cells["cutoff_low"] = writer.last_cursor

			writer.write("Obergrenze (%):", style = self.style_heading)
			writer.skip().write(float(self._exam.grading_scheme.parameters["cutoff_high"]), style = self._styles["#%"]).advance()
			self._cells["cutoff_high"] = writer.last_cursor
			writer.advance()
		else:
			raise NotImplementedError(self._exam.grading_scheme.grading_scheme_type)

		for text in [ "Gesamtzahl Arbeiten:", "Bestanden:", "Nicht bestanden:", "Knapp bestanden bei Note:", "Knapp nicht bestanden:", "Ausgezeichnet bestanden bei Note:", "Ausgezeichnet bestanden:", "Notendurchschnitt:", "Beste Note", "Schlechteste Note:" ]:
			writer.write(text, style = self.style_heading)
			if "stats_cursor" not in self._cells:
				writer.skip()
				self._cells["stats_cursor"] = writer.cursor
			writer.advance()
		writer.advance()

		writer.writerow([ "Punkte", "Ergebnis in %", "Note" ], style = self.style_heading)
		pts = float(self._exam.structure.max_points)
		while True:
			writer.write(pts, style = self._styles["#.#"])
			writer.write(self._points_to_percent_formula(writer.cursor.left), style = self._styles["#.#%"])
			writer.write(self._percent_to_grade_formula(writer.cursor.left), style = self._styles["#.#"]).advance()
			pts -= 0.5
			if pts < 0:
				break

	def _populate_results(self):
		sheet = self.sheet_results
		writer = sheet.writer()

		heading = [ "Nachname", "Vorname", "Kurs", "E-Mail", "Matrikelnummer" ]
		for task in self._exam.structure:
			heading.append(task.name)
		for task in self._exam.structure:
			heading.append(f"Punkte: {task.name}")
		heading += [ "Punkte gesamt", "Ergebnis in %", "Note" ]
		writer.writerow(heading)
		sheet.style_range(odsexport.CellRange(sheet[(0, 0)], sheet[(len(heading) - 1, 0)]), self.style_heading)
		sheet.style_column(0, odsexport.ColStyle(width = "4cm"))
		sheet.style_column(1, odsexport.ColStyle(width = "3cm"))
		sheet.style_column(3, odsexport.ColStyle(width = "7cm", hidden = True))
		for col_id in range(5 + self._exam.structure.task_count, 5 + (2 * self._exam.structure.task_count)):
			sheet.style_column(col_id, odsexport.ColStyle(hidden = True))

		for (y, entry) in enumerate(self._entries, 1):
			row = [ entry.student.last_name, entry.student.first_name, entry.student.course, entry.student.email, entry.student.student_number ]
			writer.write_many(row)

			for (task_name, contribution) in entry.grade.breakdown_by_task.items():
				writer.write(float(contribution.original_points), style = self._styles["#.#"])
			for (task_index, (task_name, contribution)) in enumerate(entry.grade.breakdown_by_task.items()):
				original_pts_cell = writer.cursor.rel(x_offset = -self._exam.structure.task_count)
				writer.write(odsexport.Formula(f"{original_pts_cell.cell_id}*{contribution.task.scalar.numerator}/{contribution.task.scalar.denominator}"), style = self._styles["#.#"])

			points_range = odsexport.CellRange(writer.cursor.left, writer.cursor.rel(x_offset = -self._exam.structure.task_count))
			writer.write(odsexport.Formula(f"SUM({points_range})"), style = self._styles["#.#"])

			points_sum_cell = writer.cursor.left
			writer.write(self._points_to_percent_formula(points_sum_cell), style = self._styles["#.#%"])

			percent_cell = writer.cursor.left
			writer.write(self._percent_to_grade_formula(percent_cell), style = self._styles["#.#"])
			if "first_grade_cell" not in self._cells:
				self._cells["first_grade_cell"] = writer.last_cursor
			writer.advance()

		writer.advance()
		for text in [ "Ø", "Ø prozentual", "Bestwertung", "Bestwertung prozentual" ]:
			writer.write(text, style = self.style_heading).advance()

		writer.cursor = sheet[(5, len(self._entries) + 2)]
		writer.mode = writer.Mode.Column
		for (x, task) in enumerate(self._exam.structure, 5):
			cell_range = odsexport.CellRange(sheet[(x, 1)], sheet[(x, len(self._entries))])
			writer.write(odsexport.Formula(f"AVERAGE({cell_range})"), style = self._styles["#.##"])
			writer.write(odsexport.Formula(f"AVERAGE({cell_range})/{task.max_points:f}"), style = self._styles["#%"])
			writer.write(odsexport.Formula(f"MAX({cell_range})"), style = self._styles["#.##"])
			writer.write(odsexport.Formula(f"MAX({cell_range})/{task.max_points:f}"), style = self._styles["#%"])
			writer.advance()

		writer.cursor = sheet[(5 + self._exam.structure.task_count * 2, len(self._entries) + 2)]

		# Total points
		cell_range = odsexport.CellRange(sheet[(writer.cursor.x, 1)], sheet[(writer.cursor.x, len(self._entries))])
		writer.write(odsexport.Formula(f"AVERAGE({cell_range})"), style = self._styles["#.##"]).skip()
		writer.write(odsexport.Formula(f"MAX({cell_range})"), style = self._styles["#.##"]).advance()

		# Result in %
		cell_range = odsexport.CellRange(sheet[(writer.cursor.x, 1)], sheet[(writer.cursor.x, len(self._entries))])
		writer.write(odsexport.Formula(f"AVERAGE({cell_range})"), style = self._styles["#%"]).skip()
		writer.write(odsexport.Formula(f"MAX({cell_range})"), style = self._styles["#%"]).advance()

		# Grade
		cell_range = odsexport.CellRange(sheet[(writer.cursor.x, 1)], sheet[(writer.cursor.x, len(self._entries))])
		writer.write(odsexport.Formula(f"AVERAGE({cell_range})"), style = self._styles["#.#"]).skip()
		writer.write(odsexport.Formula(f"MIN({cell_range})"), style = self._styles["#.#"]).advance()

		target_range = odsexport.CellRange(sheet[(0, 1)], sheet[(7 + (2 * self._exam.structure.task_count), y)])
		self._cells["conditional_format_target_range"] = target_range
		self._cells["grade_range"] = odsexport.CellRange(sheet[(7 + (2 * self._exam.structure.task_count), 1)], sheet[(7 + (2 * self._exam.structure.task_count), y)])

	def _populate_statistics(self):
		sheet = self.sheet_grade_overview
		writer = sheet.writer(self._cells["stats_cursor"])
		grade_range = self._cells["grade_range"]

		writer.write(len(self._entries)).advance()
		num_entries = writer.last_cursor
		writer.write(odsexport.Formula(f"COUNTIFS({grade_range:a};\"<=4\")")).advance()
		writer.write(odsexport.Formula(f"COUNTIFS({grade_range:a};\">4\")"), style = self._styles["failed"]).advance()
		writer.write(4.1).advance()
		self._cells["barely_passing_grade"] = writer.last_cursor
		sheet.style_row(writer.last_cursor.y, odsexport.RowStyle(hidden = True))
		writer.write(odsexport.Formula(f"COUNTIFS({grade_range:a};\">4\";{grade_range:a};\"<=\"&{writer.cursor.up})"), style = self._styles["barely_failed"]).advance()
		writer.write(1.3).advance()
		self._cells["exceptional_grade"] = writer.last_cursor
		sheet.style_row(writer.last_cursor.y, odsexport.RowStyle(hidden = True))
		writer.write(odsexport.Formula(f"COUNTIFS({grade_range:a};\"<=\"&{writer.cursor.up})"), style = self._styles["exceptional"]).advance()
		writer.write(odsexport.Formula(f"AVERAGE({grade_range:a})"), style = self._styles["#.#"]).advance()
		writer.write(odsexport.Formula(f"MIN({grade_range:a})"), style = self._styles["#.#"]).advance()
		writer.write(odsexport.Formula(f"MAX({grade_range:a})"), style = self._styles["#.#"]).advance()

		writer = sheet.writer(self._cells["stats_cursor"].right)
		writer.advance()
		writer.write(odsexport.Formula(f"{writer.cursor.left}/{num_entries}"), style = self._styles["#%"]).advance()
		writer.write(odsexport.Formula(f"{writer.cursor.left}/{num_entries}"), style = self._styles["#%"]).advance()
		writer.advance()
		writer.write(odsexport.Formula(f"{writer.cursor.left}/{num_entries}"), style = self._styles["#%"]).advance()
		writer.advance()
		writer.write(odsexport.Formula(f"{writer.cursor.left}/{num_entries}"), style = self._styles["#%"]).advance()

	def _conditional_format(self):
		sheet = self.sheet_results

		sheet.apply_conditional_format(odsexport.ConditionalFormat(target = self._cells["conditional_format_target_range"], condition_type = odsexport.ConditionType.Formula, conditions = (
			odsexport.FormatCondition(condition = f"{self._cells['first_grade_cell']:c}>{self._cells['barely_passing_grade']:acr}", cell_style = self._styles["failed"]),
			odsexport.FormatCondition(condition = f"({self._cells['first_grade_cell']:c}>4.0) AND ({self._cells['first_grade_cell']:c}<={self._cells['barely_passing_grade']:acr})", cell_style = self._styles["barely_failed"]),
			odsexport.FormatCondition(condition = f"{self._cells['first_grade_cell']:c}<={self._cells['exceptional_grade']:acr}", cell_style = self._styles["exceptional"]),
		)))


	def write(self, output_filename: str):
		self._doc = odsexport.ODSDocument()
		self._sheets["results"] = self._doc.new_sheet("Ergebnisse")
		self._sheets["grade_overview"] = self._doc.new_sheet("Notenschlüssel")
		self._populate_grade_overview()
		self._populate_results()
		self._populate_statistics()
		self._conditional_format()

		self._doc.write(output_filename)
