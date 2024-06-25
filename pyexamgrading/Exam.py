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
from .GradingScheme import GradingScheme
from .Structure import Structure
from .Student import Students
from .ExamResults import ExamResults

@dataclasses.dataclass
class ComputedGrade():
	grade: "Grade"
	next_best_grade: "Grade"
	breakdown_by_task: collections.OrderedDict
	complete_data: bool

class Exam():
	def __init__(self, name: str, date: str, lecturer: str, grading_scheme: GradingScheme, structure: Structure, students: list["Student"] | None, results: ExamResults | None):
		self._name = name
		self._date = date
		self._lecturer = lecturer
		self._grading_scheme = grading_scheme
		self._structure = structure
		self._students = students
		if self._students is None:
			self._students = Students()
		self._results = results
		if self._results is None:
			self._results = ExamResults()

	@property
	def name(self):
		return self._name

	@property
	def date(self):
		return self._date

	@property
	def lecturer(self):
		return self._lecturer

	@property
	def students(self):
		return self._students

	@students.setter
	def students(self, value: Students):
		self._students = value

	@property
	def grading_scheme(self):
		return self._grading_scheme

	@property
	def structure(self):
		return self._structure

	@property
	def results(self):
		return self._results

	def clear_results(self):
		self._results = { }

	def grade(self, student: "Student"):
		completed_tasks = self.results.get_all(student)
		exam_grade_result = self.structure.grade(completed_tasks)
		grade = self.grading_scheme.grade(exam_grade_result.total_points, self.structure.max_points)
		next_best_grade = self.grading_scheme.next_best_grade_at(exam_grade_result.total_points, self.structure.max_points, must_be_passing_grade = True)
		computed_grade = ComputedGrade(grade = grade, next_best_grade = next_best_grade, breakdown_by_task = exam_grade_result.breakdown_by_task, complete_data = all(result.missing_data == False for result in exam_grade_result.breakdown_by_task.values()))
		return computed_grade

	@classmethod
	def from_dict(cls, exam_data: dict):
		grading_scheme = GradingScheme.from_dict(exam_data["grading_scheme"])
		structure = Structure.from_dict(exam_data["structure"])
		students = Students.from_list(exam_data.get("students", [ ]))
		results = ExamResults(exam_data.get("results"))
		return cls(name = exam_data["name"], date = exam_data["date"], lecturer = exam_data["lecturer"], grading_scheme = grading_scheme, structure = structure, students = students, results = results)

	def to_dict(self):
		return collections.OrderedDict((
			("name", self._name),
			("date", self._date),
			("lecturer", self._lecturer),
			("grading_scheme", self._grading_scheme.to_dict()),
			("structure", self._structure.to_dict()),
			("students", self._students.to_list()),
			("results", self._results.to_dict()),
		))

	@classmethod
	def load_json(cls, filename: str):
		with open(filename) as f:
			exam_data = json.load(f)
		return cls.from_dict(exam_data)

	def write_json(self, filename: str):
		serialized_data = self.to_dict()
		with open(filename, "w") as f:
			json.dump(serialized_data, f, indent = "\t")
			f.write("\n")
