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
import json
import tempfile
import base64
from pyexamgrading.MultiCommand import BaseAction
from pyexamgrading.Exam import Exam
from pyexamgrading.ResultExporter import StudentResult, ResultExporter

class ActionEmail(BaseAction):
	def _filtered_students(self):
		for student in self._exam.students:
			if (self.args.search is not None) and (not student.matches(self.args.search)):
				continue
			if (self.args.filter_course is not None) and (not self.args.filter_course.lower() in student.course.lower()):
				continue
			grade = self._exam.grade(student)
#			if self.args.only_failed and grade.grade.passing:
#				continue
			yield student

	def run(self):
		if (not self.args.force) and os.path.exists(self.args.output_filename):
			raise FileExistsError(f"Refusing to overwrite: {self.args.output_filename}")

		self._exam = Exam.load_json(self.args.exam_json)
		self._entries = [ ]

		for student in self._filtered_students():
			grade = self._exam.grade(student)
			if not grade.complete_data:
				continue

			entry = StudentResult(student = student, grade = grade)
			self._entries.append(entry)

		if len(self._entries) == 0:
			print("Nothing to export: Number of students is zero.", file = sys.stderr)
			return

		exporter = ResultExporter(exam = self._exam, entries = self._entries, min_participants_stats = self._args.min_participants_stats)
		export_data = {
			"global": {
				"exam": self._exam.to_dict(),
			},
			"individual": [ ],
		}
		for entry in self._entries:
			with tempfile.NamedTemporaryFile(prefix = "pyexamgrading_", suffix = ".pdf") as tmpfile:
				exporter.export_pdf([ entry ], tmpfile.name)
				with open(tmpfile.name, "rb") as f:
					pdf_content = f.read()

			individual = {
				"student": entry.student.to_dict(),
				"result_pdf": base64.b64encode(pdf_content).decode("ascii"),
			}
			export_data["individual"].append(individual)

		with open(self._args.output_filename, "w") as f:
			json.dump(export_data, f)
