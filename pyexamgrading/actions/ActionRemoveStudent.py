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

from pyexamgrading.MultiCommand import BaseAction
from pyexamgrading.Exam import Exam

class ActionRemoveStudent(BaseAction):
	def _remove_students_without_result(self, task_name: str):
		total_count = 0
		removed_count = 0
		for student in list(self._exam.students):
			total_count += 1
			if not self._exam.results.have(student, task_name):
				removed_count += 1
				self._exam.remove_student(student)
		remaining_count = total_count - removed_count
		print(f"Removed {removed_count} of {total_count} students who do not have results for task '{task_name}' ({remaining_count} students remaining).")

	def run(self):
		self._exam = Exam.load_json(self.args.exam_json)
		if self._args.no_result_for is not None:
			self._remove_students_without_result(self._args.no_result_for)
		if self._args.commit:
			self._exam.write_json(self.args.exam_json)
		else:
			print("Dry run: results not commited to file. Rerun with --commit to write changes.")
