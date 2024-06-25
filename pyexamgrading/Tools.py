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

from .Exceptions import UnknownElementException, UndefinedElementException, DuplicateException

class Tools():
	NO_ANSWER = object()

	@classmethod
	def input_float(cls, prompt):
		while True:
			try:
				str_value = input(prompt)
				if str_value.strip() == "":
					return cls.NO_ANSWER
				elif str_value == "-":
					return None
				return float(str_value)
			except ValueError:
				pass

	@classmethod
	def plausibilize_keys(cls, test_set, optional_keys: list | None = None, mandatory_keys: list | None = None, name: str = "dictionary key"):
		test_set = set(test_set)

		permissible_keys = set()
		if optional_keys is not None:
			permissible_keys |= set(optional_keys)
		if mandatory_keys is not None:
			mandatory_keys = set(mandatory_keys)
			permissible_keys |= mandatory_keys
		excess_keys = test_set - permissible_keys
		if len(excess_keys) > 0:
			raise UnknownElementException(f"Permissible keys for {name} are {', '.join(sorted(permissible_keys))}, but unknown key(s) found: {', '.join(sorted(excess_keys))}")

		if mandatory_keys is not None:
			missing_keys = mandatory_keys - test_set
			if len(missing_keys) > 0:
				raise UndefinedElementException(f"Mandatory keys for {name} are {', '.join(sorted(mandatory_keys))}, but not included: {', '.join(sorted(missing_keys))}")

	@classmethod
	def mutex_keys(cls, test_set, mutex_keys: set, name: str = "dictionary key"):
		test_set = set(test_set)
		mutex_keys = set(mutex_keys)
		overlapping_keys = test_set & mutex_keys
		if len(overlapping_keys) > 1:
			raise DuplicateException(f"Only one {name} allowed, but {len(overlapping_keys)} mutually exclusive options given: {', '.join(sorted(overlapping_keys))}")
		elif len(overlapping_keys) == 1:
			return overlapping_keys.pop()
		else:
			return None
