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

		step_count = round((self.args.max_percentage - self.args.min_percentage) / self.args.step_percentage) + 1
		max_value = self.args.min_percentage + ((step_count - 1) * self.args.step_percentage)
		if max_value != self.args.max_percentage:
			step_count += 1

		for i in range(step_count):
			value = min(self.args.min_percentage + (i * self.args.step_percentage), self.args.max_percentage)
			grade = gs.grade(value, 100)
			if self._args.verbose < 2:
				print(f"{value:5.1f}% {grade.text}")
			else:
				print(f"{value:5.1f}% {grade.text} {value}")
