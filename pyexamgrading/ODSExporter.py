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

import datetime
import colorsys
import odsexport
from .GradingScheme import GradingSchemeType

class ODSExporter():
	@staticmethod
	def _lighten_color(hexval: str, saturation_scalar: float = 1.0, value_scalar: float = 1.0):
		def clamp(x: float):
			if x < 0:
				x = 0
			elif x > 1:
				x = 1
			return x

		assert(hexval.startswith("#"))
		(r, g, b) = (int(hexval[offset : offset + 2], 16) / 255 for offset in range(1, 7, 2))
		(h, s, v) = colorsys.rgb_to_hsv(r, g, b)
		s *= saturation_scalar
		v *= value_scalar
		(r, g, b) = colorsys.hsv_to_rgb(h, s, v)

		return f"#{round(clamp(r) * 255):02x}{round(clamp(g) * 255):02x}{round(clamp(b) * 255):02x}"

	__COLORS = {
		"red": _lighten_color("#f44336", saturation_scalar = 0.7),
		"light-red": _lighten_color("#f44336", saturation_scalar = 0.4),
		"green": "#67ac6d",
		"yellow": _lighten_color("#e49c35", value_scalar = 1.25),
	}

	def __init__(self, exam: "Exam", entries: list, stats):
		self._exam = exam
		self._entries = entries
		self._stats = stats
		self._doc = None
		self._sheets = { }
		self._styles = {
			"heading": odsexport.CellStyle(font = odsexport.Font(bold = True)),
			"heading_ralign": odsexport.CellStyle(font = odsexport.Font(bold = True), halign = odsexport.HAlign.Right),
			"heading_90deg": odsexport.CellStyle(font = odsexport.Font(bold = True), rotation_angle = 90),
			"failed": odsexport.CellStyle(background_color = self.__COLORS["red"]),
			"barely_failed": odsexport.CellStyle(background_color = self.__COLORS["light-red"]),
			"exceptional": odsexport.CellStyle(background_color = self.__COLORS["green"]),
			"bad_result": odsexport.CellStyle(background_color = self.__COLORS["red"]),
			"mediocre_result": odsexport.CellStyle(background_color = self.__COLORS["yellow"]),
			"#.#":  odsexport.CellStyle(data_style = odsexport.DataStyleNumber.fixed(1)),
			"#.##":  odsexport.CellStyle(data_style = odsexport.DataStyleNumber.fixed(2)),
			"#":  odsexport.CellStyle(data_style = odsexport.DataStyleNumber.fixed(0)),
			"#.#%":  odsexport.CellStyle(data_style = odsexport.DataStylePercent.fixed(1)),
			"#%":  odsexport.CellStyle(data_style = odsexport.DataStylePercent.fixed(0)),
			"datetime":  odsexport.CellStyle(data_style = odsexport.DataStyleDateTime.isoformat(), halign = odsexport.HAlign.Left),
		}
		self._cells = { }
		self._writers = { }

	@property
	def sheet_results(self):
		return self._sheets["results"]

	@property
	def sheet_grade_overview(self):
		return self._sheets["grade_overview"]

	@property
	def sheet_metadata(self):
		return self._sheets["metadata"]

	@property
	def style_heading(self):
		return self._styles["heading"]

	@property
	def style_heading_90deg(self):
		return self._styles["heading_90deg"]

	@property
	def style_heading_ralign(self):
		return self._styles["heading_ralign"]

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
		return odsexport.Formula(f"{points_sum_cell}/{self._cells['max_points']:a}")

	def _populate_grade_overview(self):
		sheet = self.sheet_grade_overview

		for col_no in range(3):
			sheet.style_column(col_no, odsexport.ColStyle(width = "3cm"))

		writer = sheet.writer(sheet[ 1, 0 ])
		# Grade-specific things first
		if self._exam.grading_scheme.grading_scheme_type in [ GradingSchemeType.GermanUniversityLinear, GradingSchemeType.GermanUniversityCutoff ]:
			writer.write("Bestehensgrenze (%):", style = self.style_heading_ralign)
			writer.write(float(self._exam.grading_scheme.parameters["cutoff_low"]), style = self._styles["#%"]).advance()
			self._cells["cutoff_low"] = writer.last_cursor

			writer.write("Obergrenze (%):", style = self.style_heading_ralign)
			writer.write(float(self._exam.grading_scheme.parameters["cutoff_high"]), style = self._styles["#%"]).advance()
			self._cells["cutoff_high"] = writer.last_cursor
		else:
			raise NotImplementedError(self._exam.grading_scheme.grading_scheme_type)
		writer.advance()

		# General things thereafter
		fields = [
			("exam_count", "Gesamtzahl Arbeiten:",),
			("max_points", "Gesamtzahl Punkte:"),
			("passing_points", "Bestehensgrenze (Punkte):"),
			("passing_count", "Anzahl bestanden:"),
			("failed_count", "Anzahl nicht bestanden:"),
			("barely_passing_grade", "Knapp bestanden bei Note:"),
			("barely_failed_count", "Anzahl knapp nicht bestanden:"),
			("exceptional_grade", "Ausgezeichnet bestanden bei Note:"),
			("exceptional_passed_count", "Anzahl ausgezeichnet bestanden:"),
			("grade_average", "Notendurchschnitt:"),
			("best_grade", "Beste Note:"),
			("worst_grade", "Schlechteste Note:"),
		]
		for (field_name, text) in fields:
			writer.write(text, style = self.style_heading_ralign).advance()
			self._cells[field_name] = writer.last_cursor.right
		writer.advance()

		self._cells["grade_table_start"] = writer.cursor.left


	def _populate_statistics(self):
		sheet = self.sheet_grade_overview
		grade_range = self._cells["grade_range"]

		set_values = {
			"exam_count":					(len(self._entries), None),
			"max_points":					(float(self._exam.structure.max_points), self._styles["#.#"]),
			"passing_points":				(odsexport.Formula(f"{self._cells['cutoff_low']}*{self._cells['max_points']}"), self._styles["#.#"]),
			"passing_count":				(odsexport.Formula(f"COUNTIFS({grade_range:a};\"<=4\")"), None),
			"failed_count":					(odsexport.Formula(f"COUNTIFS({grade_range:a};\">4\")"), self._styles["failed"]),
			"barely_passing_grade":			(4.1, self._styles["#.#"]),
			"barely_failed_count":			(odsexport.Formula(f"COUNTIFS({grade_range:a};\">4\";{grade_range:a};\"<=\"&{self._cells['barely_passing_grade']})"), self._styles["barely_failed"]),
			"exceptional_grade":			(1.3, self._styles["#.#"]),
			"exceptional_passed_count":		(odsexport.Formula(f"COUNTIFS({grade_range:a};\"<=\"&{self._cells['exceptional_grade']})"), self._styles["exceptional"]),
			"grade_average":				(odsexport.Formula(f"AVERAGE({grade_range:a})"), self._styles["#.#"]),
			"best_grade":					(odsexport.Formula(f"MIN({grade_range:a})"), self._styles["#.#"]),
			"worst_grade":					(odsexport.Formula(f"MAX({grade_range:a})"), self._styles["#.#"]),
		}

		for (field_name, (value, style)) in set_values.items():
			cell = self._cells[field_name].set(value)
			if style is not None:
				cell.style(style)

		for field_name in [ "barely_passing_grade", "exceptional_grade" ]:
			cell = self._cells[field_name]
			sheet.style_row(cell.y, odsexport.RowStyle(hidden = True))

		for field_name in [ "passing_count", "failed_count", "barely_failed_count", "exceptional_passed_count" ]:
			cell = self._cells[field_name].right
			cell.set(odsexport.Formula(f"{cell.left}/{self._cells['exam_count']}")).style(self._styles["#%"])

		writer = sheet.writer(start_cell = self._cells["grade_table_start"])
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

		col_ids = {
			"last_name": 0,
			"first_name": 1,
			"course": 2,
			"email": 3,
			"student_number": 4,
			"result_pts_original": 5,
			"result_pts_scaled": 5 + self._exam.structure.task_count,
			"result_pts_sum": 5 + (2 * self._exam.structure.task_count),
			"result_pts_percent": 5 + (2 * self._exam.structure.task_count) + 1,
			"grade": 5 + (2 * self._exam.structure.task_count) + 2,
			"original_grade": 5 + (2 * self._exam.structure.task_count) + 3,
			"grade_difference": 5 + (2 * self._exam.structure.task_count) + 4,
			"pts_missing_to_pass": 5 + (2 * self._exam.structure.task_count) + 5,
		}

		heading = [ "Nachname", "Vorname", "Kurs", "E-Mail", "Matrikel" ]
		for task in self._exam.structure:
			heading.append(task.name)
		for task in self._exam.structure:
			heading.append(f"Punkte: {task.name}")
		heading += [ "Punkte gesamt", "Ergebnis in %", "Note", "Ursprüngliche Note", "Notenänderung", "Punkte zu Bestehensgrenze" ]
		writer.writerow(heading)
		odsexport.CellRange(sheet[(0, 0)], sheet[(len(heading) - 1, 0)]).style(self.style_heading)
		sheet.style_column(col_ids["last_name"], odsexport.ColStyle(width = "4cm"))
		sheet.style_column(col_ids["first_name"], odsexport.ColStyle(width = "3cm"))
		sheet.style_column(col_ids["email"], odsexport.ColStyle(width = "7cm", hidden = True))
		for i in range(self._exam.structure.task_count):
			sheet.style_column(col_ids["result_pts_original"] + i, odsexport.ColStyle(width = "1.5cm"))
		for i in range(self._exam.structure.task_count):
			sheet.style_column(col_ids["result_pts_scaled"] + i, odsexport.ColStyle(width = "1.5cm", hidden = True))
		sheet.style_column(col_ids["result_pts_sum"], odsexport.ColStyle(width = "1.5cm"))
		sheet.style_column(col_ids["result_pts_percent"], odsexport.ColStyle(width = "1.5cm"))
		sheet.style_column(col_ids["grade"], odsexport.ColStyle(width = "1.5cm"))
		sheet.style_column(col_ids["original_grade"], odsexport.ColStyle(width = "1.5cm", hidden = True))
		sheet.style_column(col_ids["grade_difference"], odsexport.ColStyle(width = "1.5cm", hidden = True))
		sheet.style_column(col_ids["pts_missing_to_pass"], odsexport.ColStyle(width = "1.5cm"))

		odsexport.CellRange(sheet[(col_ids["result_pts_original"], 0)], sheet[(col_ids["pts_missing_to_pass"], 0)]).style(self.style_heading_90deg)

		for (y, entry) in enumerate(self._entries, 1):
			row = [ entry.student.last_name, entry.student.first_name, entry.student.course, entry.student.email, entry.student.student_number ]
			writer.write_many(row)

			for (task_name, contribution) in entry.grade.breakdown_by_task.items():
				if not contribution.missing_data:
					writer.write(float(contribution.original_points), style = self._styles["#.#"])
				else:
					writer.skip()
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

			# Original grade and difference. For difference, make negative
			# numbers be the negative consenquences
			writer.write(float(entry.grade.grade.value), style = self._styles["#.#"])
			writer.write(odsexport.Formula(f"{writer.last_cursor}-{writer.last_cursor.left}"), style = self._styles["#.#"])

			missing_to_pass = f"({self._cells['passing_points']:a}-{writer.cursor.rel(x_offset = -3)})"
			display_missing_to_pass = f"ROUNDUP({missing_to_pass}*2)/2"
			writer.write(odsexport.Formula(odsexport.Formula.if_then_else(if_condition = f"{missing_to_pass}>0", then_value = f"{display_missing_to_pass}", else_value = "\"\"")), style = self._styles["#.#"])
			writer.advance()

		sheet.add_data_table(odsexport.DataTable(cell_range = writer.cell_range))
		self._cells["conditional_format_target_range"] = writer.cell_range.sub_range(y_offset = 1, height = -1)
		self._cells["grade_range"] = self._cells["first_grade_cell"].make_range(height = self._cells["conditional_format_target_range"].height)

		writer.cursor = sheet[(4, len(self._entries) + 2)]
		writer.mode = writer.Mode.Column
		for text in [ "Ø:", "Ø prozentual:", "Bestwertung:", "Bestwertung prozentual:" ]:
			writer.write(text, style = self.style_heading_ralign)
		writer.advance()

		for (x, task) in enumerate(self._exam.structure, 5):
			cell_range = odsexport.CellRange(sheet[(x, 1)], sheet[(x, len(self._entries))])
			writer.write(odsexport.Formula(odsexport.Formula.average_when_have_values(cell_range, subtotal = True)), style = self._styles["#.##"])
			writer.write(odsexport.Formula(f"{writer.last_cursor}/{task.max_points:f}"), style = self._styles["#%"])
			writer.write(odsexport.Formula(odsexport.Formula.max(cell_range, subtotal = True)), style = self._styles["#.##"])
			writer.write(odsexport.Formula(f"{writer.last_cursor}/{task.max_points:f}"), style = self._styles["#%"])
			writer.advance()

		# Add conditional formatting of averages
		average_range = writer.initial_cursor.right.make_range(width = self._exam.structure.task_count, height = 2)
		first_cell = average_range.src.down
		sheet.apply_conditional_format(odsexport.ConditionalFormat(target = average_range, condition_type = odsexport.ConditionType.Formula, conditions = (
			odsexport.FormatCondition(condition = f"{first_cell:r}<0.5", cell_style = self._styles["bad_result"]),
			odsexport.FormatCondition(condition = f"{first_cell:r}<0.75", cell_style = self._styles["mediocre_result"]),
			odsexport.FormatCondition(condition = f"{first_cell:r}>0.9", cell_style = self._styles["exceptional"]),
		)))

		# Add conditional formatting of best cases
		best_range = average_range.rel(y_offset = 2)
		first_cell = first_cell.rel(y_offset = 2)
		sheet.apply_conditional_format(odsexport.ConditionalFormat(target = best_range, condition_type = odsexport.ConditionType.Formula, conditions = (
			odsexport.FormatCondition(condition = f"{first_cell:r}<0.5", cell_style = self._styles["bad_result"]),
			odsexport.FormatCondition(condition = f"{first_cell:r}<0.75", cell_style = self._styles["mediocre_result"]),
			odsexport.FormatCondition(condition = f"{first_cell:r}>0.9", cell_style = self._styles["exceptional"]),
		)))

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

	def _conditional_format(self):
		sheet = self.sheet_results

		sheet.apply_conditional_format(odsexport.ConditionalFormat(target = self._cells["conditional_format_target_range"], condition_type = odsexport.ConditionType.Formula, conditions = (
			odsexport.FormatCondition(condition = f"{self._cells['first_grade_cell']:c}>{self._cells['barely_passing_grade']:acr}", cell_style = self._styles["failed"]),
			odsexport.FormatCondition(condition = f"({self._cells['first_grade_cell']:c}>4.0) AND ({self._cells['first_grade_cell']:c}<={self._cells['barely_passing_grade']:acr})", cell_style = self._styles["barely_failed"]),
			odsexport.FormatCondition(condition = f"{self._cells['first_grade_cell']:c}<={self._cells['exceptional_grade']:acr}", cell_style = self._styles["exceptional"]),
		)))

	def _populate_metadata(self):
		sheet = self.sheet_metadata
		sheet.style_column(0, odsexport.ColStyle(width = "5cm"))
		sheet.style_column(1, odsexport.ColStyle(width = "12cm"))

		writer = sheet.writer()
		dt_format = "%Y-%m-%d %H:%M:%S"
		writer.writerow([ "Prüfungsfach:", self._exam.name ])
		writer.writerow([ "Prüfungsdatum:", self._exam.date ])
		writer.writerow([ "Prüfer:", self._exam.lecturer ])
		writer.writerow([ "Erstellungsdatum:", datetime.datetime.now() ])
		writer.last_cursor.style(self._styles["datetime"])
		writer.writerow([ "Letzte Datenänderung:", datetime.datetime.fromtimestamp(self._exam.mtime) ])
		writer.last_cursor.style(self._styles["datetime"])
		writer.cell_range.sub_range(width = 1).style(self.style_heading)


	def write(self, output_filename: str):
		self._doc = odsexport.ODSDocument()
		self._sheets["results"] = self._doc.new_sheet("Ergebnisse")
		self._sheets["grade_overview"] = self._doc.new_sheet("Notenschlüssel")
		self._sheets["metadata"] = self._doc.new_sheet("Informationen")
		self._populate_grade_overview()
		self._populate_results()
		self._populate_statistics()
		self._conditional_format()
		self._populate_metadata()

		self._doc.write(output_filename)
