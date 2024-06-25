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

import re
import sys
import fractions
from pyexamgrading.MultiCommand import BaseAction
from pyexamgrading.MoodleCSV import MoodleCSV, MoodleCSVColumn
from pyexamgrading.Exam import Exam

class ActionImport(BaseAction):
	_VALID_TASK_HEADER_RES = [
		re.compile(r"Aufgabe:\xa0(Abgabe )?(?P<task_name>.*?) \(Punkte\)"),
		re.compile(r"(?P<task_name>.*)"),
	]

	def _get_task_name(self, field_header: str):
		for regex in self._VALID_TASK_HEADER_RES:
			rematch = regex.fullmatch(field_header)
			if rematch is not None:
				return rematch.groupdict()["task_name"]
		return None

	def run(self):
		exam = Exam.load_json(self.args.exam_json)
		csv = MoodleCSV(self.args.csv_filename)
		for row in csv:
			email = row[MoodleCSVColumn.Email]
			try:
				student = exam.students.get_student_by_email(email)
			except KeyError:
				print(f"Warning: No such student with email address {email}", file = sys.stderr)
				continue

			for (field_header, value) in row.unknown_fields.items():
				task_name = self._get_task_name(field_header)
				if task_name is not None:
					if not exam.structure.has_task_with_name(task_name):
						print(f"Warning: No task in exam with name \"{task_name}\"", file = sys.stderr)
						continue

					if value == "-":
						value = 0
					else:
						value = value.replace(",", ".")

					current_result = exam.results.get(student, task_name)
					if self._args.overwrite_results or (current_result is None):
						if current_result != value:
							print(f"{student.detailed_info} setting {task_name} to {value}")
							exam.results.set(student, task_name, fractions.Fraction(value))
		exam.write_json(self.args.exam_json)
