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
