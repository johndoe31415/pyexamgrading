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

import fractions

class ExamResults():
	def __init__(self, results_by_student_number: dict | None = None):
		self._results_by_student_number = results_by_student_number
		if self._results_by_student_number is None:
			self._results_by_student_number = { }
		self._results_by_student_number = { student_number: { name: fractions.Fraction(value) for (name, value) in self._results_by_student_number[student_number].items() } for student_number in self._results_by_student_number }

	def get_all(self, student: "Student"):
		student_key = student.student_number
		if student_key not in self._results_by_student_number:
			return { }
		return self._results_by_student_number[student_key]

	def get(self, student: "Student", task_name: str):
		student_key = student.student_number
		if student_key not in self._results_by_student_number:
			return None
		return self._results_by_student_number[student_key].get(task_name)

	def set(self, student: "Student", task_name: str, value: fractions.Fraction | None):
		student_key = student.student_number
		if student_key not in self._results_by_student_number:
			self._results_by_student_number[student_key] = { }
		self._results_by_student_number[student_key][task_name] = value

	def to_dict(self):
		return { student_number: { name: str(value) for (name, value) in self._results_by_student_number[student_number].items() } for student_number in self._results_by_student_number }
