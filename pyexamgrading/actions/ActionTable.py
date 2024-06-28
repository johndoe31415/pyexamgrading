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
from pyexamgrading.GradingScheme import GradingScheme

class ActionTable(BaseAction):
	def run(self):
		gs_params = { "scheme": self.args.scheme_type.value }
		for option in self.args.option:
			(key, value) = option.split("=", maxsplit = 1)
			gs_params[key] = value
		gs = GradingScheme.from_dict(gs_params)

		end_at = self.args.end_at if (self.args.end_at is not None) else self.args.total_points

		step_count = round((end_at - self.args.start_at) / self.args.step) + 1
		max_value = self.args.start_at + ((step_count - 1) * self.args.step)
		if max_value != end_at:
			step_count += 1

		grade_sum = 0
		for i in range(step_count):
			value = min(self.args.start_at + (i * self.args.step), end_at)
			grade = gs.grade(value, self.args.total_points)
			if self._args.verbose == 0:
				print(f"{value:.1f} of {self.args.total_points:.1f}: {value / self.args.total_points * 100:5.1f}% -> {grade.text}")
			else:
				print(f"{value:.1f} of {self.args.total_points:.1f}: {value / self.args.total_points * 100:5.1f}% = {value / self.args.total_points * 100}% -> {grade.text}")

			grade_sum += grade.value

		if self._args.verbose >= 2:
			print(f"Sum of all grades: {grade_sum:.3f}")
