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

import os
import csv
import collections
import shutil
import tempfile
import fractions
import subprocess
import mako.lookup
from pyexamgrading.GradingScheme import GradingSchemeType
from pyexamgrading.WorkDir import WorkDir
from pyexamgrading.ODSExporter import ODSExporter

StudentResult = collections.namedtuple("StudentResult", [ "student", "grade" ])

class ResultExporter():
	Statistics = collections.namedtuple("Statistics", [ "average_grade", "percentile" ])

	def __init__(self, exam: "Exam", entries: list[StudentResult], min_participants_stats: int):
		self._exam = exam
		self._entries = entries
		self._min_participants_stats = min_participants_stats
		self._stats = self._compute_stats()

	@property
	def total_student_count(self):
		return len(self._entries)

	@property
	def average_grade(self):
		return sum(entry.grade.grade.value for entry in self._entries) / self.total_student_count

	@property
	def average_points(self):
		return sum(entry.grade.grade.achieved_points for entry in self._entries) / self.total_student_count

	def _compute_stats(self):
		grades = [ ]
		for entry in self._entries:
			grades.append((entry.grade.grade.value, entry.grade.grade.text))
		grades.sort()

		last_value = None
		percentile = { }
		for (counter, (value, text)) in enumerate(grades):
			if value != last_value:
				last_value = value
				percentile[text] = 100 * (len(self._entries) - counter) / len(self._entries)

		return self.Statistics(average_grade = self.average_grade, percentile = percentile)

	def export_csv(self, entries: list[StudentResult], filename: str):
		self._entries.sort(key = lambda entry: (entry.student.course, entry.student.last_name, entry.student.first_name))
		with open(filename, "w") as f:
			writer = csv.writer(f)

			heading = [ "Kurs", "Nachname", "Vorname", "E-Mail", "Matrikelnummer" ]
			for task in self._exam.structure:
				heading.append(task.name)
			heading += [ "Punkte (gewichtet verrechnet)", "Gesamtpunkte", "Ergebnis in %", "Note" ]
			writer.writerow(heading)

			for entry in entries:
				student = entry.student
				row = [ student.course, student.last_name, student.first_name, student.email, student.student_number ]
				for (name, contribution) in entry.grade.breakdown_by_task.items():
					row.append(float(contribution.original_points))
				row.append(float(entry.grade.grade.achieved_points))
				row.append(float(entry.grade.grade.max_points))
				row.append(float(entry.grade.grade.achieved_points / entry.grade.grade.max_points * 100))
				row.append(entry.grade.grade.text)
				writer.writerow(row)

	def export_all_csv(self, filename: str):
		return self.export_csv(entries = self._entries, filename = filename)

	def export_tex(self, entries: list[StudentResult], filename: str):
		def error_function(msg):
			raise Exception(msg)
		self._entries.sort(key = lambda entry: (entry.student.course, entry.student.student_number))
		lookup = mako.lookup.TemplateLookup([ f"{os.path.dirname(__file__)}/templates" ],strict_undefined = True)
		template = lookup.get_template("export.tex")
		template_vars = {
			"entries": entries,
			"exam": self._exam,
			"stats": self._stats,
			"min_participants_stats": self._min_participants_stats,
			"error": error_function,

			"fractions": fractions,
			"GradingSchemeType": GradingSchemeType,
		}
		rendered = template.render(**template_vars)
		with open(filename, "w") as f:
			f.write(rendered)

	def export_all_tex(self, filename: str):
		return self.export_tex(entries = self._entries, filename = filename)

	def export_pdf(self, entries: list[StudentResult], filename: str):
		with tempfile.TemporaryDirectory() as tmpdir:
			tex_filename = f"{tmpdir}/pyexam.tex"
			pdf_filename = f"{tmpdir}/pyexam.pdf"
			self.export_tex(entries = entries, filename = tex_filename)
			with WorkDir(tmpdir):
				subprocess.check_call([ "pdflatex", tex_filename ])
			shutil.move(pdf_filename, filename)

	def export_all_pdf(self, filename: str):
		return self.export_pdf(entries = self._entries, filename = filename)

	def export_all_ods(self, filename: str):
		ods_exporter = ODSExporter(exam = self._exam, entries = self._entries, stats = self._stats)
		ods_exporter.write(filename)
