import os
from pyexamgrading.MultiCommand import BaseAction
from pyexamgrading.Exam import Exam
from pyexamgrading.Tools import Tools

class ActionEnterResults(BaseAction):
	def run(self):
		exam = Exam.load_json(self.args.exam_json)
		first = True
		while True:
			if first:
				first = False
			else:
				print()
			search_key = input("Student search key: ")
			if search_key == "":
				continue
			student = exam.students.search(search_key)
			if len(student) == 0:
				print("No student found matching this key.")
				continue
			elif len(student) > 1:
				print(f"Ambiguous key found, {len(student)} results:")
				for finding in student:
					print(f"	{finding.detailed_info}")
				continue

			student = student.pop()
			print(f"Entering data for: {student.detailed_info}")

			for task in exam.structure:
				current_result = exam.results.get(student, task.name)
				if (not self._args.enter_all_results) and (current_result is not None):
					continue
				if current_result is None:
					current_result_str = "currently no result"
				else:
					current_result_str = f"currently: {current_result}"

				new_result = Tools.input_float(f"{task.name} (max. {task.max_points:.1f} pts, {current_result_str}): ")
				if new_result != Tools.NO_ANSWER:
					exam.results.set(student, task.name, float(new_result))
					exam.write_json(self.args.exam_json)
