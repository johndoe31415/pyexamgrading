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

import enum
import collections
import fractions
from .Tools import Tools

class GradingSchemeType(enum.Enum):
	GermanUniversityLinear = "german-university-linear"
	GermanUniversityCutoff = "german-university-cutoff"

class GradingScheme():
	Grade = collections.namedtuple("Grade", [ "text", "value", "passing", "achieved_points", "max_points" ])
	HypotheticalGrade = collections.namedtuple("HypotheticalGrade", [ "point_difference", "grade" ])

	def __init__(self, grading_scheme_type: GradingSchemeType, parameters: dict):
		self._grading_scheme_type = grading_scheme_type
		self._parameters = parameters

	@property
	def grading_scheme_type(self):
		return self._grading_scheme_type

	@property
	def parameters(self):
		return self._parameters

	@classmethod
	def from_dict(cls, data: dict):
		grading_scheme_type = GradingSchemeType(data["scheme"])
		parameters = { }
		match grading_scheme_type:
			case GradingSchemeType.GermanUniversityLinear | GradingSchemeType.GermanUniversityCutoff:
				Tools.plausibilize_keys(data, optional_keys = [ "scheme", "cutoff_low", "cutoff_high" ], name = "German university grading scheme")
				parameters["cutoff_low"] = fractions.Fraction(data.get("cutoff_low", 50)) / 100
				parameters["cutoff_high"] = fractions.Fraction(data.get("cutoff_high", 100)) / 100
		return cls(grading_scheme_type = grading_scheme_type, parameters = parameters)

	def grade(self, points: float, max_points: float):
		ratio = points / max_points
		match self.grading_scheme_type:
			case GradingSchemeType.GermanUniversityLinear | GradingSchemeType.GermanUniversityCutoff:
				cutoff_range = self._parameters["cutoff_high"] - self._parameters["cutoff_low"]
				posrange = (self._parameters["cutoff_high"] - ratio) / cutoff_range
				computed_grade = 1 + (3 * posrange)

				if computed_grade > fractions.Fraction(4):
					# Non-passing grade
					passing = False
					computed_grade = max(fractions.Fraction("4.1"), computed_grade)
				else:
					passing = True

				cutoff_grade = fractions.Fraction(5) if (self.grading_scheme_type == GradingSchemeType.GermanUniversityLinear) else fractions.Fraction(4)
				if computed_grade > cutoff_grade:
					computed_grade = fractions.Fraction(5)
				elif computed_grade < fractions.Fraction(1):
					computed_grade = fractions.Fraction(1)

				rounded_grade = round(computed_grade, 1)
				text = f"{rounded_grade:.1f}"
				return self.Grade(text = text, value = rounded_grade, passing = passing, achieved_points = points, max_points = max_points)

	def next_best_grade_at(self, points: float, max_points: float, step: float = 0.5, max_steps: int = 100, must_be_passing_grade: bool = False):
		base_grade = self.grade(points, max_points).text
		for i in range(1, max_steps + 1):
			point_difference = i * step
			hypothetical_points = points + point_difference
			hypothetical_grade = self.grade(hypothetical_points, max_points)
			if (hypothetical_grade.text != base_grade) and ((not must_be_passing_grade) or hypothetical_grade.passing):
				return self.HypotheticalGrade(point_difference = point_difference, grade = hypothetical_grade)
		return None

	def to_dict(self):
		result = collections.OrderedDict((
			("scheme", self.grading_scheme_type.value),
		))
		match self.grading_scheme_type:
			case GradingSchemeType.GermanUniversityLinear | GradingSchemeType.GermanUniversityCutoff:
				result["cutoff_low"] = str(self._parameters["cutoff_low"] * 100)
				result["cutoff_high"] = str(self._parameters["cutoff_high"] * 100)
		return result

	def __str__(self):
		match self.grading_scheme_type:
			case GradingSchemeType.GermanUniversityLinear | GradingSchemeType.GermanUniversityCutoff:
				return f"{self.grading_scheme_type.name} ({self._parameters['cutoff_low'] * 100:.0f}% to {self._parameters['cutoff_high'] * 100:.0f}%)"
