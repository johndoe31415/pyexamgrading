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

import re
import zipfile
import collections
import dataclasses

@dataclasses.dataclass
class SubmittedFile():
	student_name: str
	submission_id: str
	filename: str
	zip_fileinfo: zipfile.ZipInfo
	zip_file: zipfile.ZipFile

	def read(self) -> bytes:
		return self.zip_file.read(self.zip_fileinfo.filename)

class MoodleZIP():
	_ZIPFILE_FILENAME_REGEX = re.compile(r"(?P<student_name>[^_]+)_(?P<submission_id>\d+)_assignsubmission_file/(?P<filename>.*)")

	def __init__(self, filename: str):
		self._filename = filename
		self._zipfile = None
		self._files = None

	def _parse_directories(self):
		self._files = collections.defaultdict(list)
		for zip_fileinfo in self._zipfile.filelist:
			rematch = self._ZIPFILE_FILENAME_REGEX.fullmatch(zip_fileinfo.filename)
			if rematch is None:
				continue
			submitted_file = SubmittedFile(student_name = rematch["student_name"], submission_id = rematch["submission_id"], filename = rematch["filename"], zip_fileinfo = zip_fileinfo, zip_file = self._zipfile)
			self._files[submitted_file.submission_id].append(submitted_file)

	def __enter__(self):
		if self._zipfile is not None:
			raise ValueError(f"ZIP-file {self._filename} already open.")
		self._zipfile = zipfile.ZipFile(self._filename, "r")
		self._parse_directories()
		return self

	def __exit__(self, *args):
		self._zipfile.close()

	def get_files_for_submission_id(self, submission_id: str):
		return self._files[submission_id]

	def get_file_for_submission_id(self, submission_id: str):
		files = self.get_files_for_submission_id(submission_id)
		if len(files) != 1:
			raise ValueError(f"Trying to get a single unique fiel for submission ID {submission_id}, but {len(files)} submissions found for that ID.")
		return files.pop()

if __name__ == "__main__":
	with MoodleZIP("/tmp/TINF_KryptoOffSecRevEng-ed25519 Public Keys-353121.zip") as mzip:
		subfile = mzip.get_file_for_submission_id("909290")
		print(subfile)
		print(subfile.read())
