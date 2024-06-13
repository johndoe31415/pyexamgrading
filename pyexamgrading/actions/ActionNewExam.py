import os
from pyexamgrading.MultiCommand import BaseAction
from pyexamgrading.Student import Students
from pyexamgrading.Exam import Exam

class ActionNewExam(BaseAction):
	def run(self):
		if (not self.args.force) and os.path.exists(self.args.exam_json):
			raise FileExistsError(f"Refusing to overwrite: {self.args.exam_json}")
		exam = Exam.load_json(self.args.exam_definition_json)

		for student_filename in self.args.students_json:
			students = Students.load_students_json(student_filename)
			exam.students.add_all_active(students)

		exam.write_json(self.args.exam_json)
