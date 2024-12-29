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
import sys
from pyexamgrading.MultiCommand import BaseAction
from pyexamgrading.Exam import Exam
from pyexamgrading.ResultExporter import StudentResult, ResultExporter

class ActionExport(BaseAction):
	def _filtered_students(self):
		for student in self._exam.students:
			if (self.args.search is not None) and (not student.matches(self.args.search)):
				continue
			if (self.args.filter_course is not None) and (not self.args.filter_course.lower() in student.course.lower()):
				continue
			yield student

	@property
	def file_output_type(self):
		if self.args.output_type == "auto":
			(prefix, suffix) = os.path.splitext(self.args.output_filename)
			if suffix not in [ ".tex", ".csv", ".pdf", ".ods" ]:
				raise ValueError("Filename must end in a known extension.")
			return suffix[1:]
		else:
			return self.args.output_type

	def run(self):
		if (not self.args.force) and os.path.exists(self.args.output_filename):
			raise FileExistsError(f"Refusing to overwrite: {self.args.output_filename}")

		self._exam = Exam.load_json(self.args.exam_json)
		self._entries = [ ]

		for student in self._filtered_students():
			grade = self._exam.grade(student)
			if (not grade.complete_data) and (not self.args.show_all):
				continue

			entry = StudentResult(student = student, grade = grade)
			self._entries.append(entry)

		if len(self._entries) == 0:
			print("Nothing to export: Number of students is zero.", file = sys.stderr)
			return

		exporter = ResultExporter(exam = self._exam, entries = self._entries, min_participants_stats = self._args.min_participants_stats)
		export_handler = getattr(exporter, f"export_all_{self.file_output_type}")
		export_handler(self.args.output_filename)
