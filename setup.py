import setuptools

with open("README.md") as f:
	long_description = f.read()

setuptools.setup(
	name = "pyexamgrading",
	packages = setuptools.find_packages(),
	version = "0.0.1",
	license = "gpl-3.0",
	description = "Manage grade computation of university exams",
	long_description = long_description,
	long_description_content_type = "text/markdown",
	author = "Johannes Bauer",
	author_email = "joe@johannes-bauer.com",
	url = "https://github.com/johndoe31415/pyexamgrading",
	download_url = "https://github.com/johndoe31415/pyexamgrading/archive/0.0.1.tar.gz",
	keywords = [ "python", "exam", "university" ],
	install_requires = [
		"mako",
	],
	entry_points = {
		"console_scripts": [
			"pyexam = pyexamgrading.__main__:main"
		]
	},
	include_package_data = True,
	classifiers = [
		"Development Status :: 5 - Production/Stable",
		"Intended Audience :: Developers",
		"License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
		"Programming Language :: Python :: 3",
		"Programming Language :: Python :: 3 :: Only",
		"Programming Language :: Python :: 3.12",
	],
)
