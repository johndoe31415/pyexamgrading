import fractions

class ExamResults():
	def __init__(self, results_by_student_number: dict | None):
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
