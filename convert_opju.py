import PyOrigin,os,atexit,hashlib
import originpro as op
import pygetwindow as gw
import pandas as pd
import numpy as np
from pathlib import Path
from PIL import Image
from datetime import datetime

#%%####### README

# To use this script:
#
# Open an .opju file in Origin 2024b
# Go to Window > Script Window
# In the Script Window, run the following command, where you
# replace (script) with the full path of this file:
#
# run -pyf (script)

#%%####### functions to navigate and communicate with LabTalk

def make_note_window():
	"""This function creates a note window for communication between
	Python and LabTalk within Origin. To reduce the chance of
	conflicting with an existing note window, the window name is
	randomly generated.
	
	Returns:
		str: Origin note window name		
	"""

	ok = False
	
	# try to create a window name based on random hex string
	# this approach is not immune to race condition
	# so there could be conflicts if such window was
	# created after python got settled on the filename
	while not ok:
		window_name = "Note_{0}".format(os.urandom(3).hex())
		exists = op.find_notes(window_name)
		if exists is None:
			ok = True
									
	# create note window
	PyOrigin.LT_execute(f"window -n n {window_name}")

	return window_name

def wrap_LTexecute(window_name, LT_cmd):
	"""This function is a wrapper to execute a LabTalk command and
	retrieve the result to Python. The LabTalk output is redirected
	to a note window, where it can be read by Python. The note window
	is destroyed after the LabTalk output is read.
	
	Args:
		window_name (str): name of the Origin note window
		LT_cmd (str): LabTalk command
	
	Returns:
		str: the output of the LabTalk command	
	"""
	
	# command to redirect LabTalk output to note window
	start_redirect = f"type.notes$={window_name}; type.redirection=2"
	
	# command to end redirection
	end_redirect = "type.redirection=32768"

	# execute LabTalk command
	# this also activates the note window
	cmd = f"{start_redirect}; {LT_cmd}; {end_redirect}"
	PyOrigin.LT_execute(cmd)

	# get note page object
	Notes=PyOrigin.ActiveNotePage()
	
	# get text
	text = Notes.GetText()

	# destroy note window
	Notes.Destroy()
	
	return text

def get_files_folders(opju_folder):
	"""This function retrieves a list of files and folders in the
	given directory. It runs a LabTalk command "pe_dir oname:=objects"
	and retrieves the result via a note window. The note window is
	used since PyOrigin does not have a method or function to list
	subfolders in the current folder, only files.
	
	Args:
		opju_folder (str): Origin folder path
		
	Returns:
		list: A list containing file names
		list: A list containing folder names	
	"""

	op.pe.cd(opju_folder)

	# create Note window
	# this is used for communication between python and LabTalk
	window_name = make_note_window()
	
	# for some reason the oname is enough to output folder contents
	# also tried to add "type objects$" but that duplicated the output
	cmd = "pe_dir oname:=objects"
	paths = wrap_LTexecute(window_name, cmd)
	
	# clean result
	paths = [p.rstrip("\r") for p in paths.split("\n")]
	paths = [p for p in paths if p not in ["",window_name]]
	
	# get file names
	files = [p for p in paths if not p.startswith("<Folder>")]
	
	# get and clean folder names
	folders = [p.split("> ",1)[1] for p in paths if p.startswith("<Folder>")]
	folders = [f"{opju_folder}{f}/" for f in folders]
	
	return files, folders
	
def traverse_folder(project_path,opju_folder):
	"""This function traverses the files and folders in an .opju file
	in a recursive manner.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path	
	"""
	
	files, folders = get_files_folders(opju_folder)

	handle_files(project_path,opju_folder)
	
	for folder in folders:
		traverse_folder(project_path,folder)

def check_folder(filepath):
	"""This function creates the folders for a given file path if
	the folders do not exist.
	
	Args:
		filepath (str): full file path
	"""
	
	folder = filepath.rsplit("/",1)[0] + "/"
	Path(folder).mkdir(parents=True, exist_ok=True)

def print_and_log(f,string):
	"""This function prints a message to stdout and writes it to the 
	given opened file object.
	
	Args:
		f (_io.TextIOWrapper): an opened file object
		string (str): A message to print and log
	"""
	
	print(string)
	f.write(string + "\n")

#%%####### functions to handle and write content types

def file_path(filename):
	"""This function converts any slash characters (/, \) in 
	a filename will to underscore (_).
	
	"""
	return filename.replace("/","_").replace("\\","_")
	
def out_path(project_path,opju_folder,filename):
	"""This function makes a full file path given the args.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path
		filename (str): filename

	Returns:
		str: full file path
	"""
	
	return f"{project_path}{opju_folder}{filename}"

def report_save(ftype, path):
	"""This function outputs a message indicating that a file of
	the given type was saved to a path.
	
	Args:
		ftype (str): file type
		path (str): full file path
	"""
	
	# do not print preceeding slash if exists
	if path[0] == "/" and path[1] != "/":
		path = path[1:]
	
	print_and_log(f," {0} | saved to: {1}".format(ftype.ljust(6, " "),path))

def handle_files(project_path,opju_folder):
	""" This function is a master handling function for all files
	in the current Origin working directory.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path	
	"""

	# PGTYPES of PyOrigin
	# PyOrigin.PGTYPE_GRAPH 3 <- graph
	# PyOrigin.PGTYPE_IMAGE 29 <- image
	# PyOrigin.PGTYPE_LAYOUT 11 <- layout
	# PyOrigin.PGTYPE_MATRIX 5 <- matrix
	# PyOrigin.PGTYPE_NOTES 9 <- notes
	# PyOrigin.PGTYPE_WKS 2 <- worksheet
	
	# get Folder object for current location
	Folder = PyOrigin.ActiveFolder()
	
	# get items in the Origin project
	for page in Folder.PageBases():

		t =  page.GetType()

		# worksheets
		if t == PyOrigin.PGTYPE_WKS:
			book = op.find_book("w", page.Name)
			handle_book(project_path,opju_folder,book,"w")
		
		# matrices
		elif t == PyOrigin.PGTYPE_MATRIX:
			book = op.find_book("m", page.Name)
			handle_book(project_path,opju_folder,book,"m")

		# graphs
		elif t == PyOrigin.PGTYPE_GRAPH:
			graph = op.find_graph(page.Name)
			handle_graph(project_path,opju_folder,graph)

		# images
		elif t == PyOrigin.PGTYPE_IMAGE:
			image = op.find_image(page.Name)
			handle_image(project_path,opju_folder,image)
			
		# notes
		elif t == PyOrigin.PGTYPE_NOTES:
			notes = op.find_notes(page.Name)
			handle_notes(project_path,opju_folder,notes)

		# unimplemented type
		else:
			print_and_log(f,"UNIMPLEMENTED TYPE:", page, type(page))

def handle_book(project_path, opju_folder, book, btype):
	"""This function handles each worksheet in an Origin workbook.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path
		btype (str): "w" for workbook or "m" for matrix
	"""
	
	index = 0
	end = False
	
	# have not so far found a more clean way to determine
	# how many worksheets are in a book
	while not end:
		try:
			sheet = book.__getitem__(index)
			
			if btype == "w": 
				handle_worksheet(project_path,opju_folder,sheet,book.name)
			elif btype == "m":
				handle_matrixsheet(project_path,opju_folder,sheet,book.name)
			index += 1
		except TypeError:
			end = True

def handle_worksheet(project_path,opju_folder,sheet,book_name):
	"""This function handles an Origin worksheet and saves its contents
	to a .csv file.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path
		sheet (originpro.worksheet.WSheet): Origin worksheet object
		book_name (str): parent workbook name
	"""
	
	ncols = sheet.obj.GetColumns().GetCount()
	
	df = sheet.to_df()				
	
	units = [sheet.obj[n].GetUnits() for n in range(ncols)]
	comments = [sheet.obj[n].GetComments() for n in range(ncols)]
	formulas = [sheet.obj[n].GetFormula() for n in range(ncols)]
	
	fname = file_path(f"{book_name}.{sheet.name}.csv")
	out = out_path(project_path,opju_folder,fname)
	check_folder(out)

	# write out first with this
	# this produces a good csv file
	df.to_csv(out, encoding="utf-8")

	# read the file and add additional headers
	lines = []
	with open(out, "r", encoding="utf-8") as f:
		header = f.readline().rstrip("\n").rstrip("\r")
		lines.append(header)
		for label,lst in [["Units",units],["Comments",comments],["F(x)=",formulas]]:
			tmp = [l for l in lst if l != ""]
			if len(tmp) > 0:
				lines.append(",".join([label] + [str(l) for l in lst]))
	
		for line in f:
			line = line.rstrip("\n").rstrip("\r")
			lines.append(line)

	# write to file
	with open(out, "w", encoding="utf-8") as f:
		for line in lines:
			f.write(line + "\n")

	report_save("book",f"{opju_folder}{fname}")

def handle_matrixsheet(project_path,opju_folder,matrixsheet,matrix_name):
	"""This function saves each matrix object in a matrixsheet to a
	.csv file.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path
		matrixsheet (originpro.matrix.MSheet): an Origin matrixsheet object
		matrix_name (str): parent matrix name
	"""
	
	# handle each matrix object separately
	for m in matrixsheet.obj.GetMatrixObjects():
		
		# it seems that units can't be accessed
		# although they can be set in Origin
		# could not figure out how to get them with Python

		data = m.GetData()
		longname = m.GetLongName()
		if longname == "":
			longname = f"object{m.Name}"

		fname = file_path(f"{matrix.name}.{longname}.csv")
		out = out_path(project_path,opju_folder,fname)
		check_folder(out)
		
		pd.DataFrame(data).to_csv(out, encoding="utf-8")
		
		report_save("matrix",f"{opju_folder}{fname}")

def handle_graph(project_path,opju_folder,graph):
	"""This function saves an Origin graph to a .png file.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path
		graph (originpro.graph.GPage): an Origin graph object
	"""
	
	fname = file_path(f"{graph.name}.png")
	out = out_path(project_path,opju_folder,fname)
	check_folder(out)
	
	graph.save_fig(out,type="png", replace=True, width=1920)
	report_save("graph",f"{opju_folder}{fname}")

def handle_image(project_path,opju_folder,image):
	"""This function saves an Origin image to a .png file.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path
		image (originpro.image.IPage): an Origin image object
	"""

	width, height = image.size
	argb_array = image.to_np()

	# convert ARGB to RGBA
	# ARGB format: 0xAARRGGBB
	# RGBA format: 0xRRGGBBAA
	rgba_array = np.empty((height, width, 4), dtype=np.uint8)
	rgba_array[..., 0] = (argb_array >> 16) & 0xFF  # red
	rgba_array[..., 1] = (argb_array >> 8) & 0xFF   # green
	rgba_array[..., 2] = argb_array & 0xFF          # blue
	rgba_array[..., 3] = (argb_array >> 24) & 0xFF  # alpha

	fname = file_path(f"{opju_folder}{image.name}.png")
	out = out_path(project_path,opju_folder,fname)
	check_folder(out)

	Image.fromarray(rgba_array, "RGBA").save(out)
	report_save("image",f"{opju_folder}{fname}")

def handle_notes(project_path,opju_folder,notes):
	"""This function saves the content of an Origin notes window
	to a .txt or a .html file depending on the Origin syntax of
	the notes.
	
	Args:
		project_path (str): full file path of the .opju file
		opju_folder (str): Origin folder path
		notes (originpro.notes.Notes): an Origin notes object
	"""
	
	out = out_path(project_path,opju_folder,notes.name)
	
	# notes syntax
	notes_stx = notes.syntax

	# normal text or origin rich text
	if notes_stx in [0,3]:
		suffix = "txt"
	# html
	elif notes_stx == 1:
		suffix = "html"
	# markdown
	elif notes_stx == 2:
		suffix = "md"
	
	fname = file_path(f"{notes.name}.{suffix}")
	out = out_path(project_path,opju_folder,fname)
	
	check_folder(out)
	with open(out, "w", encoding="utf-8") as f:
		f.write(notes.text)
		
	report_save("notes",out)

#%%####### process file

# get Origin window title
window = [w for w in gw.getAllTitles() if "Origin 202" in w]
if len(window) == 0:
	print("\nERROR: could not find Origin window title")
	exit(1)
else:
	window = window[0]

# get project name and project path from window title	
file_name, project_path, _ = window.split(" - ",2)

# clean project path
project_path = Path(project_path).as_posix()

# clean project name
file_name = file_name.rstrip("*").strip()

# calculate file hash
with open(f"{project_path}/{file_name}", "rb") as f:
    file_hash = hashlib.md5()
    chunk = f.read(8192)
    while chunk:
        file_hash.update(chunk)
        chunk = f.read(8192)
file_hash = file_hash.hexdigest()

# clean project name more
project_name = file_name.rsplit(".",1)[0]

# go to root and store path in variable
opju_root = op.pe.cd("/")

out_dir = f"{project_path}/{project_name}"
proceed = True

# check if output file already exists
# if exists, prompt user to select how to proceed
if os.path.isdir(out_dir):
	warning = "\n".join(["WARNING: folder already exists:",
						out_dir,
						"overwrite files? (y/n)"])
	answer = input(warning).lower()
	while answer not in ["yes","y","no","n"]:
		answer = input(warning).lower()
	if answer in ["no","n"]:
		proceed = False
else:
	Path(out_dir).mkdir(parents=True, exist_ok=True)

if proceed:
	
	# log file
	log_out = f"{out_dir}/convert.log"

	with open(log_out, "w", encoding="utf-8") as f:
		atexit.register(f.close)
		
		now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
		print_and_log(f,f"started run at:\n{now}\n")
		
		print_and_log(f, f"process file:\n{file_name}")
		print_and_log(f, f"\nfile md5sum:\n{file_hash}\n")
		
		print_and_log(f, "handling files...")
		# recursively process all subfolders and .opju objects
		traverse_folder(project_path, opju_root)
		print_and_log(f, "done")
		
		print_and_log(f, "\nsave log to:\nconvert.log")

print("\nfinished run\n")