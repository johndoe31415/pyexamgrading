import os
import csv
import fractions
import collections
import mako.lookup
from pyexamgrading.MultiCommand import BaseAction
from pyexamgrading.Exam import Exam
from pyexamgrading.Tools import Tools
from pyexamgrading.GradingScheme import GradingSchemeType

class ActionExport(BaseAction):
	DisplayEntry = collections.namedtuple("DisplayEntry", [ "student", "grade" ])
	Statistics = collections.namedtuple("Statistics", [ "average_grade", "percentile" ])

	@property
	def total_student_count(self):
		return len(self._entries)

	@property
	def passed_student_count(self):
		return self._counts["passed_students"]

	@property
	def failed_student_count(self):
		return self.total_student_count - self.passed_student_count

	@property
	def incomplete_data_count(self):
		return self._counts["incomplete_data"]

	@property
	def average_grade(self):
		return sum(entry.grade.grade.value for entry in self._entries) / self.total_student_count

	@property
	def average_points(self):
		return sum(entry.grade.grade.achieved_points for entry in self._entries) / self.total_student_count

	def _filtered_students(self):
		for student in self._exam.students:
			if (self.args.search is not None) and (not student.matches(self.args.search)):
				continue
			if (self.args.filter_course is not None) and (not self.args.filter_course.lower() in student.course.lower()):
				continue
			yield student

	def _export_csv(self):
		self._entries.sort(key = lambda entry: (entry.student.course, entry.student.last_name, entry.student.first_name))
		with open(self.args.output_filename, "w") as f:
			writer = csv.writer(f)

			heading = [ "Kurs", "Nachname", "Vorname", "E-Mail", "Matrikelnummer" ]
			for task in self._exam.structure:
				heading.append(task.name)
			heading += [ "Punkte (gewichtet verrechnet)", "Gesamtpunkte", "Ergebnis in %", "Note" ]
			writer.writerow(heading)

			for entry in self._entries:
				student = entry.student
				row = [ student.course, student.last_name, student.first_name, student.email, student.student_number ]
				for (name, contribution) in entry.grade.breakdown_by_task.items():
					row.append(float(contribution.original_points))
				row.append(float(entry.grade.grade.achieved_points))
				row.append(float(entry.grade.grade.max_points))
				row.append(float(entry.grade.grade.achieved_points / entry.grade.grade.max_points * 100))
				row.append(entry.grade.grade.text)
				writer.writerow(row)

	def _export_tex(self):
		def error_function(msg):
			raise Exception(msg)
		self._entries.sort(key = lambda entry: entry.student.student_number)
		lookup = mako.lookup.TemplateLookup([ f"{os.path.dirname(__file__)}/../templates" ],strict_undefined = True)
		template = lookup.get_template("export.tex")
		template_vars = {
			"entries": self._entries,
			"exam": self._exam,
			"stats": self._stats,
			"min_participants_stats": self.args.min_participants_stats,
			"error": error_function,

			"fractions": fractions,
			"GradingSchemeType": GradingSchemeType,
		}
		rendered = template.render(**template_vars)
		with open(self.args.output_filename, "w") as f:
			f.write(rendered)

	def _compute_stats(self):
		grades = [ ]
		for entry in self._entries:
			grades.append((entry.grade.grade.value, entry.grade.grade.text))
		grades.sort()

		last_value = None
		percentile = { }
		for (counter, (value, text)) in enumerate(grades):
			if value != last_value:
				last_value = value
				percentile[text] = 100 * (len(self._entries) - counter) / len(self._entries)

		average_grade = sum(entry.grade.grade.value for entry in self._entries) / len(self._entries)
		return self.Statistics(average_grade = average_grade, percentile = percentile)

	@property
	def file_output_type(self):
		if self.args.output_type == "auto":
			(prefix, suffix) = os.path.splitext(self.args.output_filename)
			if suffix not in [ ".tex", ".csv" ]:
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
			if not grade.complete_data:
				continue

			entry = self.DisplayEntry(student = student, grade = grade)
			self._entries.append(entry)

		self._stats = self._compute_stats()
		export_handler = getattr(self, f"_export_{self.file_output_type}")
		export_handler()
