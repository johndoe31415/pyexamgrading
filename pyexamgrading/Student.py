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

import json
import dataclasses
import collections

from .Exceptions import DuplicateException

@dataclasses.dataclass(order = True)
class Student():
	last_name: str
	first_name: str
	email: str
	student_number: str
	course: str
	active: bool = True

	@property
	def full_name(self):
		return f"{self.last_name}, {self.first_name}"

	@property
	def detailed_info(self):
		return f"{self.full_name} <{self.email}> ({self.course}, {self.student_number})"

	@classmethod
	def from_dict(cls, student_dict: dict):
		return cls(last_name = student_dict["last_name"], first_name = student_dict["first_name"], email = student_dict["email"], student_number = student_dict["student_number"], course = student_dict.get("course"), active = student_dict.get("active", True))

	def to_dict(self):
		result = dataclasses.asdict(self, dict_factory = collections.OrderedDict)
		if self.active:
			del result["active"]
		return result

	def matches(self, search_key: str):
		search_key = search_key.lower()
		if self.student_number.lower().startswith(search_key):
			return True
		if search_key in self.email.lower():
			return True
		if search_key in self.last_name.lower():
			return True
		return False

class Students():
	def __init__(self):
		self._students_by_email = { }
		self._students_by_student_number = { }
		self._active_student_count = 0

	@property
	def active_student_count(self):
		return self._active_student_count

	def add(self, student: Student):
		if student.email in self._students_by_email:
			raise DuplicateException("Cannot add student {student}, duplicate email: {self._students_by_email[student.email]} already present in database.")
		if student.student_number in self._students_by_student_number:
			raise DuplicateException("Cannot add student {student}, duplicate student number: {self._students_by_student_number[student.student_number]} already present in database.")
		self._students_by_email[student.email] = student
		self._students_by_student_number[student.student_number] = student
		if student.active:
			self._active_student_count += 1
		return student

	@classmethod
	def load_students_json(cls, filename: str):
		with open(filename) as f:
			student_list = json.load(f)
		return cls.from_list(student_list)

	@classmethod
	def from_list(cls, student_list: list[dict]):
		students = cls()
		for student_dict in student_list:
			student = Student.from_dict(student_dict)
			students.add(student)
		return students

	def search(self, key):
		return [ student for student in self if student.matches(key) ]

	def to_list(self):
		return [ student.to_dict() for student in sorted(self) ]

	def get_student_by_email(self, email: str):
		return self._students_by_email[email]

	def get_student_by_student_number(self, student_number: str):
		return self._students_by_student_number[student_number]

	def add_all_active(self, students: "Students"):
		for student in students:
			if student.active:
				self.add(student)

	def add_all(self, students: "Students"):
		for student in students:
			self.add(student)

	def __iter__(self):
		return iter(self._students_by_email.values())

	def __len__(self):
		return len(self._students_by_email)

	def __str__(self):
		return f"Students <{len(self)} students, {self.active_student_count} active>"
