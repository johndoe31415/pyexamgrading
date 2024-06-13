# pyexamgrading
This is a collection of tools I'm using for grading university exams.

## Example
There are several anonymized files in the `test_files` directory to get you
started. Some have the exact format of a MOODLE document. First, you need to
create an exam JSON from a list of students who participated and a course
structure definition (which indicates how everything will be graded). Note that
you can have several course files which you all import:

```
$ pyexam new test_files/exam_definition_kryptologie.json graded.json test_files/course_99cs*
```

You can then start importing data from MOODLE (e.g., laboratory exercises). In
the given example exam, laboratory exercises are weighted 25% and final exam is
weighted 75%.

```
$ pyexam import graded.json test_files/moodle_export.csv
```

We can now view results but final grades will not be shown because data is
incomplete so far (only lab exercises has been imported):

```
$ pyexam print graded.json

3 students omitted with incomplete data.
```

We can force those to print:

```
$ pyexam print -a graded.json
⚠ 99CS1  Asmussen, Fips                           5.0      0.7 / 80.0 (0.9%)        39.5 pts missing for 4.0
⚠ 99CS2  Bar, Berta                               5.0      10.7 / 80.0 (13.4%)      29.5 pts missing for 4.0
⚠ 99CS2  Foo, Aaron                               5.0      11.4 / 80.0 (14.2%)      29.0 pts missing for 4.0

0 of 3 pass (0.0%), 3 failed (100.0%)
Average grade: 5.0 (7.6 points)

Grade histogram:
1.0 - 1.5: 0 (0.0%)
1.5 - 2.0: 0 (0.0%)
2.0 - 2.5: 0 (0.0%)
2.5 - 3.0: 0 (0.0%)
3.0 - 3.5: 0 (0.0%)
3.5 - 4.0: 0 (0.0%)
4.0 - 4.5: 0 (0.0%)
4.5 - 5.0: 3 (100.0%)     ********************************************************************************

Point histogram:
10.2 - 11.4: 2 (66.7%)    ********************************************************************************
9.1 - 10.2: 0 (0.0%)
8.0 - 9.1: 0 (0.0%)
6.8 - 8.0: 0 (0.0%)
5.7 - 6.8: 0 (0.0%)
4.6 - 5.7: 0 (0.0%)
3.4 - 4.6: 0 (0.0%)
2.3 - 3.4: 0 (0.0%)
1.1 - 2.3: 0 (0.0%)
0.0 - 1.1: 1 (33.3%)      ****************************************
```

Note that the exlamation mark in the beginning of each line means not all data
has been imported. We can also search for a particular student by name/student
number and get a detailed breakdown of their performance:

```
$ pyexam print -a -s asmussen -b graded.json
⚠ 99CS1  Asmussen, Fips                           5.0      0.7 / 80.0 (0.9%)        39.5 pts missing for 4.0
        • Klausur Aufgabe 1 0.0 / 10.0 (0.0%) -> 0.000 pts
        • Klausur Aufgabe 2 0.0 / 10.0 (0.0%) -> 0.000 pts
        • Klausur Aufgabe 3 0.0 / 10.0 (0.0%) -> 0.000 pts
        • Klausur Aufgabe 4 0.0 / 10.0 (0.0%) -> 0.000 pts
        • Klausur Aufgabe 5 0.0 / 10.0 (0.0%) -> 0.000 pts
        • Klausur Aufgabe 6 0.0 / 10.0 (0.0%) -> 0.000 pts
        • Labor 1 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 1 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 2 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 2 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 3 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 3 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 4 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 4 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 5 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 5 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 6 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 6 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 7 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 7 / Teil 2 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 8 / Teil 1 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 8 / Teil 2 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 9 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 9 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
```

Internally, all values are computed as exact fractions, which we can also show:

```
$ pyexam print -a -s asmussen -b -vv graded.json
⚠ 99CS1  Asmussen, Fips                           5.0      25/36 / 80.0 (0.9%)      39.5 pts missing for 4.0
        • Klausur Aufgabe 1 0.0 / 10.0 (0.0%) -> 0.0% * 1 = 0 pts
        • Klausur Aufgabe 2 0.0 / 10.0 (0.0%) -> 0.0% * 1 = 0 pts
        • Klausur Aufgabe 3 0.0 / 10.0 (0.0%) -> 0.0% * 1 = 0 pts
        • Klausur Aufgabe 4 0.0 / 10.0 (0.0%) -> 0.0% * 1 = 0 pts
        • Klausur Aufgabe 5 0.0 / 10.0 (0.0%) -> 0.0% * 1 = 0 pts
        • Klausur Aufgabe 6 0.0 / 10.0 (0.0%) -> 0.0% * 1 = 0 pts
        • Labor 1 / Teil 1 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 1 / Teil 2 0.0 / 8.0 (0.0%) -> 0.0% * 5/36 = 0 pts
        • Labor 2 / Teil 1 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 2 / Teil 2 0.0 / 8.0 (0.0%) -> 0.0% * 5/36 = 0 pts
        • Labor 3 / Teil 1 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 3 / Teil 2 0.0 / 8.0 (0.0%) -> 0.0% * 5/36 = 0 pts
        • Labor 4 / Teil 1 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 4 / Teil 2 0.0 / 8.0 (0.0%) -> 0.0% * 5/36 = 0 pts
        • Labor 5 / Teil 1 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 5 / Teil 2 0.0 / 8.0 (0.0%) -> 0.0% * 5/36 = 0 pts
        • Labor 6 / Teil 1 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 6 / Teil 2 0.0 / 8.0 (0.0%) -> 0.0% * 5/36 = 0 pts
        • Labor 7 / Teil 1 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 7 / Teil 2 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 8 / Teil 1 0.0 / 8.0 (0.0%) -> 0.0% * 5/36 = 0 pts
        • Labor 8 / Teil 2 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 9 / Teil 1 0.5 / 8.0 (6.2%) -> 6.2% * 5/36 = 5/72 pts
        • Labor 9 / Teil 2 0.0 / 8.0 (0.0%) -> 0.0% * 5/36 = 0 pts
```

We can make a hypothethis how the grades will look when the remaining data has
the same average value, best-case or worst-case values:

```
$ pyexam print -a -s asmussen -b -H avg graded.json
    99CS1  Asmussen, Fips                           5.0      2.8 / 80.0 (3.5%)        37.5 pts missing for 4.0
        • Klausur Aufgabe 1 0.3 / 10.0 (3.5%) -> 0.347 pts
        • Klausur Aufgabe 2 0.3 / 10.0 (3.5%) -> 0.347 pts
        • Klausur Aufgabe 3 0.3 / 10.0 (3.5%) -> 0.347 pts
        • Klausur Aufgabe 4 0.3 / 10.0 (3.5%) -> 0.347 pts
        • Klausur Aufgabe 5 0.3 / 10.0 (3.5%) -> 0.347 pts
        • Klausur Aufgabe 6 0.3 / 10.0 (3.5%) -> 0.347 pts
        • Labor 1 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 1 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 2 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 2 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 3 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 3 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 4 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 4 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 5 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 5 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 6 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 6 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 7 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 7 / Teil 2 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 8 / Teil 1 0.0 / 8.0 (0.0%) -> 0.000 pts
        • Labor 8 / Teil 2 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 9 / Teil 1 0.5 / 8.0 (6.2%) -> 0.069 pts
        • Labor 9 / Teil 2 0.0 / 8.0 (0.0%) -> 0.000 pts
[...]
⚠ Shown grades are hypothetical according to avg model. ⚠
```

Let us import the remainder of the data. This is easiest by replicating the CSV
format of MOODLE:

```
$ pyexam import graded.json test_files/final_exam.csv
```

Now we can print the final results:

```
$ pyexam print graded.json
    99CS1  Asmussen, Fips                           4.4      35.1 / 80.0 (43.9%)      5.0 pts missing for 4.0
    99CS2  Bar, Berta                               5.0      19.7 / 80.0 (24.6%)      20.5 pts missing for 4.0
    99CS2  Foo, Aaron                               2.8      56.4 / 80.0 (70.5%)      0.5 pts missing for 2.7

1 of 3 pass (33.3%), 2 failed (66.7%)
Average grade: 4.1 (37.1 points)

Grade histogram:
1.0 - 1.5: 0 (0.0%)
1.5 - 2.0: 0 (0.0%)
2.0 - 2.5: 0 (0.0%)
2.5 - 3.0: 1 (33.3%)      ********************************************************************************
3.0 - 3.5: 0 (0.0%)
3.5 - 4.0: 0 (0.0%)
4.0 - 4.5: 1 (33.3%)      ********************************************************************************
4.5 - 5.0: 1 (33.3%)      ********************************************************************************

Point histogram:
50.8 - 56.4: 1 (33.3%)    ********************************************************************************
45.1 - 50.8: 0 (0.0%)
39.5 - 45.1: 0 (0.0%)
33.8 - 39.5: 1 (33.3%)    ********************************************************************************
28.2 - 33.8: 0 (0.0%)
22.6 - 28.2: 0 (0.0%)
16.9 - 22.6: 1 (33.3%)    ********************************************************************************
11.3 - 16.9: 0 (0.0%)
5.6 - 11.3: 0 (0.0%)
0.0 - 5.6: 0 (0.0%)
```

If you want to enter data manually, you can use the enter mode, which will ask
you for a query key first (student number or parts of their last name). The
program will refuse to use ambiguous search keys (e.g., if two students have
the same name or their student number starts with the same digits):

```
$ pyexam enter graded.json
Student search key: asmus
Entering data for: Asmussen, Fips <s333333@student.dhbw-mannheim.de> (99CS1, 543890433)
Klausur Aufgabe 6 (max. 10.0 pts, currently no result): 4
Labor 9 / Teil 1 (max. 8.0 pts, currently no result): 7
Labor 9 / Teil 2 (max. 8.0 pts, currently no result): 6

Student search key:
```

Note also that the program will only query for results that are not yet entered.

Finally, once you are finished with data entry you can generate a TeX file that
is individual to each student and which shows a detailed breakdown of their
grade. This can be printed and attached to the finals, for example.

## License
GNU GPL-3.
