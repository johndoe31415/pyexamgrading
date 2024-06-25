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

import collections
from pyexamgrading.MultiCommand import BaseAction
from pyexamgrading.Exam import Exam

class ColorScheme():
	FgRed = "\x1b[31m"
	FgBlack = "\x1b[30m"
	FgYellow = "\x1b[33m"
	BgYellow = "\x1b[43m"
	Normal = "\x1b[0m"

	StudentFailed = FgRed
	StudentPassed = Normal
	WarningMsg = BgYellow + FgBlack

class ActionPrint(BaseAction):
	DisplayEntry = collections.namedtuple("DisplayEntry", [ "student", "grade" ])

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

	def _print_entry(self, entry):
		if entry.grade.complete_data:
			indicator = ""
		else:
			indicator = f"{self.color.WarningMsg}⚠{self.color.Normal}"

		if entry.grade.next_best_grade is None:
			next_best_grade_str = ""
		else:
			next_best_grade_str = f"{entry.grade.next_best_grade.point_difference:.1f} pts missing for {entry.grade.next_best_grade.grade.text}"

		if self._args.verbose < 2:
			result_str = f"{entry.grade.grade.achieved_points:.1f} / {entry.grade.grade.max_points:.1f} ({entry.grade.grade.achieved_points / entry.grade.grade.max_points * 100:.1f}%)"
		else:
			result_str = f"{entry.grade.grade.achieved_points} / {entry.grade.grade.max_points:.1f} ({entry.grade.grade.achieved_points / entry.grade.grade.max_points * 100:.1f}%)"
		color = self.color.StudentPassed if entry.grade.grade.passing else self.color.StudentFailed
		print(f"{indicator:<3s} {color}{entry.student.course:<6s} {entry.student.full_name:<40s} {entry.grade.grade.text:<8s} {result_str:<24s} {next_best_grade_str}{self.color.Normal}")
		if self.args.breakdown:
			for (task_name, task_breakdown) in entry.grade.breakdown_by_task.items():
				if self._args.verbose < 2:
					equivalent_pts_str = f"{task_breakdown.scaled_points:.3f} pts"
				else:
					equivalent_pts_str = f"{task_breakdown.original_points / task_breakdown.task.max_points * 100:.1f}% * {task_breakdown.task.scalar} = {task_breakdown.scaled_points} pts"
				print(f"        • {task_name} {task_breakdown.original_points:.1f} / {task_breakdown.task.max_points:.1f} ({task_breakdown.original_points / task_breakdown.task.max_points * 100:.1f}%) -> {equivalent_pts_str}")
			print()

	def _print_summary(self):
		if self.incomplete_data_count > 0:
			print(f"{self.incomplete_data_count} students omitted with incomplete data.")

		if self.total_student_count > 0:
			print(f"{self.passed_student_count} of {self.total_student_count} pass ({self.passed_student_count / self.total_student_count * 100:.1f}%), {self.failed_student_count} failed ({self.failed_student_count / self.total_student_count * 100:.1f}%)")
			print(f"Average grade: {self.average_grade:.1f} ({self.average_points:.1f} points)")
		print()

	def _sort_entries(self):
		match self.args.sort_criteria:
			case "name":
				self._entries.sort(key = lambda entry: (entry.student.last_name, entry.student.first_name))

			case "grade":
				self._entries.sort(key = lambda entry: (-entry.grade.grade.achieved_points, entry.student.last_name, entry.student.first_name))

	def _print_histogram(self, name: str, bucket_count: int, min_value: float, max_value: float, get_value: "callable", reverse: bool = False):
		if self.total_student_count == 0:
			return

		bucket_width = (max_value - min_value) / bucket_count
		buckets = [ 0 ] * bucket_count
		if bucket_width == 0:
			return

		print(f"{name} histogram:")
		for entry in self._entries:
			bucket_index = int((get_value(entry) - min_value) / bucket_width)
			if bucket_index < 0:
				bucket_index = 0
			elif bucket_index >= bucket_count:
				bucket_index = bucket_count - 1
			buckets[bucket_index] += 1

		max_bar_width = 80
		scale = max_bar_width / max(buckets)
		order = enumerate(buckets) if (not reverse) else reversed(list(enumerate(buckets)))
		for (bucket_index, count) in order:
			low = min_value + bucket_index * bucket_width
			high = min_value + (bucket_index + 1) * bucket_width
			lstr = f"{low:.1f} - {high:.1f}: {count} ({count / self.total_student_count * 100:.1f}%)"
			bar_width = round(count * scale)
			bar = "*" * bar_width
			print(f"{lstr:<25s} {bar}")
		print()

	def _print_histograms(self):
		if self.total_student_count == 0:
			return

		self._print_histogram("Grade", 8, 1.0, 5.0, get_value = lambda entry: entry.grade.grade.value)

		max_points = max(entry.grade.grade.achieved_points for entry in self._entries)
		self._print_histogram("Point", 10, 0, max_points, get_value = lambda entry: entry.grade.grade.achieved_points, reverse = True)

	def run(self):
		self.color = ColorScheme()
		self._exam = Exam.load_json(self.args.exam_json)
		self._entries = [ ]
		self._counts = {
			"incomplete_data": 0,
			"passed_students": 0,
		}

		for student in self._filtered_students():
			grade = self._exam.grade(student)
			if self.args.hypothesize != "no":
				match self.args.hypothesize:
					case "avg":
						finished_total_points = sum(contribution.task.max_points for contribution in grade.breakdown_by_task.values() if not contribution.missing_data)
						finished_achieved_points = sum(contribution.original_points for contribution in grade.breakdown_by_task.values() if not contribution.missing_data)
						achieved_ratio = finished_achieved_points / finished_total_points if (finished_total_points != 0) else 0

					case "best":
						achieved_ratio = 1

					case "half":
						achieved_ratio = 0.5

					case "worst":
						achieved_ratio = 0

				for contribution in grade.breakdown_by_task.values():
					if not contribution.missing_data:
						continue
					self._exam.results.set(student, contribution.task.name, achieved_ratio * contribution.task.max_points)

				# Re-grade hypothetical grade
				grade = self._exam.grade(student)

			entry = self.DisplayEntry(student = student, grade = grade)

			if (not entry.grade.complete_data) and (not self.args.show_all):
				self._counts["incomplete_data"] += 1
				continue

			if entry.grade.grade.passing:
				self._counts["passed_students"] += 1
			self._entries.append(entry)

		self._sort_entries()
		for entry in self._entries:
			self._print_entry(entry)
		print()
		self._print_summary()
		self._print_histograms()

		if self.args.hypothesize != "no":
			indicator = "⚠"
			print(f"{indicator} Shown grades are hypothetical according to {self.args.hypothesize} model. {indicator}")
