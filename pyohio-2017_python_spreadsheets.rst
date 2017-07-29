Python and Spreadsheets: State of the Union, July 2017
=====================================================
Kojo Idrissa
-------------
-  QA Specialist for Decisio Health
-  Former Educator & Accountant w/ an MBA
-  `@transition <https://twitter.com/transition>`_


Outline
========
-  Demo some fundamental OpenPyXL data types and how it deals with spreadsheets
-  Demo some basic Python tasks with spreadsheets
-  Discuss some sticking points when trying to work with spreadsheets via code


Setup
=======
-  Not a lot of advanced code here
-  Very beginner-friendly
-  Interesting code provided by *your* specific use case

Context/Background
==================
-  My role as a Python contributor
    *  Grow the community
    *  Bring non-programmers into the fold
-  Solutions for people who aren't professional developers
    *  Solutions obvious to experienced developers often aren't available to non-developers

Two primary audiences
=====================
-  Spreadsheet users who want to "step their game up"

-  Python developers confronted with spreadsheets

Three Basic Approaches
=======================

-  Working with spreadsheet files (my focus)
-  Running Python code in a spreadsheet app: our Lightning Talk on `xlwings <https://www.xlwings.org/>`_
-  Scripting spreadsheet apps: theoretically possible with OpenOffice & LibreOffice, `just not easy <https://onesheep.org/scripting-libreoffice-python/>`_


Working with spreadsheet files
===============================
-  `Python-Excel.org <http://www.python-excel.org>`_ web site collects info for working with Excel files
-  Use case: You've been given a spreadsheet full of data to process, but not in a way that's convenient in your spreadsheet app.
    +  Perfect entry point for creating new Pythonistas. Spreadsheet hackers are ALMOST programmers *(kinda)*.
-  `OpenPyXL <http://pythonhosted.org//openpyxl/>`_
    +  get it from `PyPi <https://pypi.python.org/pypi/openpyxl>`_ 
    +  Python >=2.6
    +  works with .xlsx & .xlsm files
+  Let's look at some examples!

Beginner Demo
==============
-  reading from a spreadsheet of well-structured data (timesheet data)

	Lots of business applications are just an RDBMS (relational database) with a GUI. So, they often "export" database tables. Those look exactly like spreadsheets.

-  writing to a spreadsheet (timesheet aggregation results)

	What if you've got code that generates a lot of data that you want/need to share with someone? And that someone (who's NOT a developer) want's to be able to manipulate it in a spreadsheet?

Advanced Beginner Demo
=========================
-  Convert a spreadsheet into JSON 
	-  Pass the utilization data to another program or store it.
	-  (KPop Draft Config File)
	
		What if you have an app that lets you build your own KPop group in a fantasy draft, to see how "popular" that group is? It's like a KPop Entertainment Company Simulator. This does NOT exist, but I MIGHT build it.

Caveats
========

Spreadsheets are often used as a visual medium
-----------------------------------------------
-  What if your input data isn't well-structured?
-  What if you're output has to be "visual"?