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
import dataclasses
import collections
import functools
import fractions
from .Tools import Tools
from .Exceptions import DuplicateException

class ScalingMode(enum.Enum):
	Points = "scale_points"
	Percentage = "scale_percentage"

@dataclasses.dataclass
class StructureTask():
	name: str
	max_points: fractions.Fraction
	group: str = "default"
	scalar: fractions.Fraction | None = None
	bonus: bool = False
	scale_mode: ScalingMode | None = None
	scale_value: fractions.Fraction | None = None

	def to_dict(self):
		result = collections.OrderedDict((
			("name", self.name),
			("max_points", str(self.max_points)),
		))
		if self.group != "default":
			result["group"] = self.group
		if self.bonus:
			result["bonus"] = True
		if self.scale_mode is not None:
			result[self.scale_mode.value] = str(self.scale_value)
		return result


class Structure():
	TaskContribution = collections.namedtuple("TaskContribution", [ "task", "original_points", "scaled_points", "missing_data" ])
	ExamGradeResult = collections.namedtuple("ExamGradeResult", [ "total_points", "breakdown_by_task" ])

	def __init__(self, reference_group: str | None, groups: dict[str, float] | None):
		self._reference_group = reference_group or "default"
		self._groups = groups
		if self._groups is None:
			self._groups = { }
		self._tasks_by_name = collections.OrderedDict()

	@functools.cached_property
	def max_points_reference_group(self):
		return sum(task.max_points for task in self if (not task.bonus) and (task.group == self._reference_group))

	@functools.cached_property
	def max_points(self):
		return sum(task.scalar * task.max_points for task in self if (not task.bonus))

	@functools.cached_property
	def reference_group_weight(self):
		return self._groups[self._reference_group]

	@functools.cached_property
	def max_points_by_group(self):
		group_weight = collections.defaultdict(fractions.Fraction)
		for task in self:
			group_weight[task.group] += task.max_points
		return group_weight

	def has_task_with_name(self, task_name: str):
		return task_name in self._tasks_by_name

	def add_task(self, task: StructureTask):
		if self.has_task_with_name(task.name):
			raise DuplicateException(f"Task with name '{task.name}' already exists.")
		self._tasks_by_name[task.name] = task
		return self

	def grade(self, completed_tasks: dict[str, fractions.Fraction]):
		total_points = fractions.Fraction(0)
		breakdown_by_task = collections.OrderedDict()
		for task in self:
			original_points = completed_tasks.get(task.name)
			if original_points is None:
				missing_data = True
				original_points = 0
			else:
				missing_data = False
			scaled_points = task.scalar * original_points
			total_points += scaled_points
			breakdown_by_task[task.name] = self.TaskContribution(task = task, original_points = original_points, scaled_points = scaled_points, missing_data = missing_data)
		return self.ExamGradeResult(total_points = total_points, breakdown_by_task = breakdown_by_task)

	def to_dict(self):
		return {
			"groups": { key: str(value) for (key, value) in self._groups.items() },
			"reference_group": self._reference_group,
			"tasks": [ task.to_dict() for task in self ],
		}

	def dump(self):
		for task in self:
			print(task)
		print()

	def _compute_scalars(self):
		reference_group_max_points = self.max_points_by_group[self._reference_group]
		for task in self:
			if task.group == self._reference_group:
				task.scalar = 1
			else:
				relative_scalar = self._groups[task.group] / self.reference_group_weight * reference_group_max_points / self.max_points_by_group[task.group]
				task.scalar = relative_scalar

			match task.scale_mode:
				case ScalingMode.Points:
					task.scalar = task.scalar * task.scale_value / task.max_points

				case ScalingMode.Percentage:
					task.scalar = task.scalar * (task.scale_value / 100) / task.max_points * self.max_points

	@classmethod
	def from_dict(cls, data: dict):
		Tools.plausibilize_keys(data, optional_keys = [ "groups", "reference_group" ], mandatory_keys = [ "tasks" ])
		groups = { key: fractions.Fraction(value) for (key, value) in data["groups"].items() } if ("groups" in data) else None
		structure = cls(reference_group = data.get("reference_group"), groups = groups)
		for task_data in data["tasks"]:
			Tools.plausibilize_keys(task_data, optional_keys = [ "bonus", "group", "scale_points", "scale_percentage" ], mandatory_keys = [ "name", "max_points" ])
			scale_mode = Tools.mutex_keys(task_data, set([ "scale_points", "scale_percentage" ]))
			if scale_mode is not None:
				scale_value = fractions.Fraction(task_data[scale_mode])
				scale_mode = ScalingMode(scale_mode)
			else:
				scale_value = None
			task = StructureTask(name = task_data["name"], max_points = fractions.Fraction(task_data["max_points"]), group = task_data.get("group", "default"), bonus = task_data.get("bonus", False), scale_mode = scale_mode, scale_value = scale_value)
			structure.add_task(task)
		structure._compute_scalars()
		return structure

	def __iter__(self):
		return iter(self._tasks_by_name.values())
