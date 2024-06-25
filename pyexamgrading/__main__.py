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

import sys
import pyexamgrading
from .MultiCommand import MultiCommand
from .actions.ActionNewExam import ActionNewExam
from .actions.ActionEnterResults import ActionEnterResults
from .actions.ActionImport import ActionImport
from .actions.ActionPrint import ActionPrint
from .actions.ActionExport import ActionExport

def main():
	mc = MultiCommand(description = "Grade exams and allow for import and export of various data", trailing_text = f"pyexamgrading v{pyexamgrading.VERSION}", run_method = True)

	def genparser(parser):
		parser.add_argument("-f", "--force", action = "store_true", help = "Overwrite output file if it already exists.")
		parser.add_argument("-v", "--verbose", action = "count", default = 0, help = "Increase verbosity. Can be given multiple times.")
		parser.add_argument("exam_definition_json", help = "Input JSON filename containing the exam stucture.")
		parser.add_argument("exam_json", help = "Output JSON filename containing the graded exam.")
		parser.add_argument("students_json", nargs = "+", help = "Input JSON filename(s) containing the participants of the exam.")
	mc.register("new", "Create a new exam file", genparser, action = ActionNewExam)

	def genparser(parser):
		parser.add_argument("-a", "--enter-all-results", action = "store_true", help = "Ask for input of all results, even if they have been already entered.")
		parser.add_argument("-v", "--verbose", action = "count", default = 0, help = "Increase verbosity. Can be given multiple times.")
		parser.add_argument("exam_json", help = "JSON filename containing the graded exam.")
	mc.register("enter", "Interactively enter graded data", genparser, action = ActionEnterResults)

	def genparser(parser):
		parser.add_argument("-o", "--overwrite-results", action = "store_true", help = "Overwrite results when they are already entered.")
		parser.add_argument("-v", "--verbose", action = "count", default = 0, help = "Increase verbosity. Can be given multiple times.")
		parser.add_argument("exam_json", help = "JSON filename containing the graded exam.")
		parser.add_argument("csv_filename", help = "CSV file to import results from.")
	mc.register("import", "Import CSV data, for example from MOODLE", genparser, action = ActionImport)

	def genparser(parser):
		parser.add_argument("-a", "--show-all", action = "store_true", help = "Show all students, even those with incomplete data.")
		parser.add_argument("-s", "--search", metavar = "pattern", help = "Show only students which match this pattern.")
		parser.add_argument("-c", "--filter-course", metavar = "pattern", help = "Show only students which match this course.")
		parser.add_argument("-S", "--sort-criteria", choices = [ "name", "grade" ], default = "name", help = "Sort shown students by these criteria. Can be one of %(choices)s, defaults to %(default)s.")
		parser.add_argument("-b", "--breakdown", action = "store_true", help = "Show an individual task breakdown for each result.")
		parser.add_argument("-H", "--hypothesize", choices = [ "no", "best", "half", "worst", "avg" ], default = "no", help = "When not all grades are present, model a grade hypothesis. Can be one of %(choices)s, defaults to '%(default)s'.")
		parser.add_argument("-v", "--verbose", action = "count", default = 0, help = "Increase verbosity. Can be given multiple times.")
		parser.add_argument("exam_json", help = "JSON filename containing the graded exam.")
	mc.register("print", "Show exam data and grading", genparser, action = ActionPrint)

	def genparser(parser):
		parser.add_argument("-t", "--output-type", choices = [ "auto", "csv", "tex", "pdf" ], default = "auto", help = "Export in this output format. Can be one of %(default)s, defaults to %(default)s. When 'auto', the filename extension must clearly indicate the file type.")
		parser.add_argument("-s", "--search", metavar = "pattern", help = "Show only students which match this pattern.")
		parser.add_argument("-c", "--filter-course", metavar = "pattern", help = "Show only students which match this course.")
		parser.add_argument("--min-participants-stats", metavar = "count", type = int, default = 10, help = "By default, statistical information is not shown for privacy purposes below this number of participants of a test. By default, this is %(default)d.")
		parser.add_argument("-f", "--force", action = "store_true", help = "Overwrite output file if it already exists.")
		parser.add_argument("-v", "--verbose", action = "count", default = 0, help = "Increase verbosity. Can be given multiple times.")
		parser.add_argument("exam_json", help = "JSON filename containing the graded exam.")
		parser.add_argument("output_filename", help = "Output filename containing the rendered data.")
	mc.register("export", "Export exam data", genparser, action = ActionExport)

	returncode = mc.run(sys.argv[1:])
	return returncode or 0

if __name__ == "__main__":
	sys.exit(main())
