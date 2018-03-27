from pathlib import Path
from time import sleep as s
from csv import DictReader as dr
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from os import remove as deletefile

import subprocess as sp
import cmd

from numpy import arange as even_ticks
import matplotlib.pyplot as plt

def file_picker():  
    s(1)  
    root = Tk()
    filename = askopenfilename()
    root.withdraw()
    return Path(filename)


class Report:
	def __init__(self, in_csv):
		self._csv = in_csv
		with open(self._csv, newline='') as csvfile:
			srcdat = dr(csvfile)
			self._dict = [line for line in srcdat]

	def __iter__(self):
		for line in self._dict:
			yield line

	def __len__(self):
		return len(self._dict)

	def _make_chart(self, row, datalables):
		"""
		datalables must be tuple
		"""
		datalist = [float(self._dict[row][i]) 
					for i in datalables]
		student = self._dict[row]['name']
		x_pos = even_ticks(len(datalables))
		plt.bar(x_pos, datalist, align='center', color='c', alpha=0.5)
		plt.xticks(x_pos, datalables)
		plt.ylabel('achivement')
		plt.yticks(self.yscale)
		plt.title(f'progress for {student}')
		plt.savefig(f'{student}.png')
		plt.clf()

	def _make_HTML(self, row, commentlist):
		studentname = self._dict[row]['name']
		start = f"""
			<h1> Student Name : {studentname}</h1>
			<img src='{studentname}.png' />
		"""

		for comment in commentlist:
			start += f"""
				<h2>{comment}</h2>
				<p>{self._dict[row][comment]}</p>
			"""

		with open(f'{studentname}.html', 'w') as html:
			html.write(start)

	def _make_docx(self, row):
			studentname = self._dict[row]['name']
			arg_list = ['pandoc',
				f'{studentname}.html',
				f'-o {studentname}.docx']
			p = sp.Popen(arg_list, shell=True)

	def _cleanup(self, row):
		studentname = self._dict[row]['name']
		deletefile(f"{studentname}.html")
		deletefile(f"{studentname}.png")

	@property
	def keys(self):
		return [key for key
			in self._dict[0]]

	def set_scale(self, ymin, ymax, ystep):
		self.yscale = even_ticks(ymin, ymax, ystep)

	def make(self, chartdata, htmlcomments):
		rows = range(len(self._dict)) 
		for row in rows:
			self._make_chart(row, chartdata)
			self._make_HTML(row, htmlcomments)
			self._make_docx(row)

	def clean(self):
		rows = range(len(self._dict)) 
		for row in rows:
			self._cleanup(row)


class Report_cmd(cmd.Cmd):

	prompt = "Report Maker > "
	chartkeys = []
	commentkeys = []

	def _check_csv(self):
		if not 'name' in self._report.keys:
			print("your source file lacks a 'name' column")
			print("this is the only required label")
		print("column headings:")
		print(self._report.keys)

	def do_getdata(self, arg):
		"select the source csv file"
		try:
			self._report = Report(arg)
			self._check_csv()
		except FileNotFoundError:
			try:
				self._report = Report(file_picker())
				self._check_csv()

			except ValueError:
				print("opps, something went wrong!")

	def _chartkey(self, arg):
		args = arg.split(" ")
		for key in args:
			if key in self._report.keys:
				self.chartkeys.append(key)

	def _showkeys(self):
		print("chart keys selected:")
		print(self.chartkeys)
		print("comment keys selected:")
		print(self.commentkeys)

	def do_chartkey(self, arg):
		"select the chart keys"
		self._chartkey(arg)
		self._showkeys()

	def do_quickkey(self, arg):
		"select chartkeys, remaining keys are assumed to be commentkeys"
		self._chartkey(arg)
		for key in self._report.keys:
			if key not in self.chartkeys:
				self.commentkeys.append(key)
		if "name" in self.commentkeys:
			self.commentkeys.remove("name")
		self._showkeys()

	def do_clearkeys(self, arg):
		"uh ho you added a key by mistake"
		self.chartkeys = []
		self.commentkeys = []
		self._showkeys()

	def do_comment(self, arg):
		"select the text comment headings"
		args = arg.split(" ")
		for key in args:
			if key in self._report.keys:
				self.commentkeys.append(key)
		self._showkeys()

	def do_make(self, arg):
		"""
		takes the upper scale parameter for the y axis
		as an argument. e.g. 'make 100' for percent
		Then creates the png, html, and docx files"""
		try:
			upper = float(arg)
			test = {
			lambda u : u > 100: lambda u : self._report.set_scale(0, u, 25),
			lambda u : u > 50: lambda u : self._report.set_scale(0, u, 10),
			lambda u : u > 10: lambda u : self._report.set_scale(0, u, 5),
			lambda u : True: lambda u : self._report.set_scale(0, u, 2)
			}
			for key in test:
				if key(upper):
					test[key](upper)
					break
		
			print("upper y limit set to : ", upper)
			try:
				self._report.make(self.chartkeys, self.commentkeys)
				print(" files created with no obvious errors")
			except ValueError:
				print("opps something went wrong")
		except ValueError:
			print("you need to set the upper limit for your y axis")

	def do_cleanup(self, arg):
		"clean up the intermediate files, (optional)"
		self._report.clean()

	def do_q(self, arg):
		"q is for quitters"
		return True

if __name__ == '__main__':
	Report_cmd().cmdloop()
