import os
import sys
import shutil
import Tkinter
import tkFileDialog
import json
from docx import *
import zipfile
from zipfile import ZipFile
import time
import glob
import tkinter as tk
import ntpath
import win32print



try:
	commandKey = sys.argv[1]
	
	#prevent file from some characters like :| symbols
	
	def preventexcelrename(file_string):
		sheet_name_out1s = list(file_string)
		sheetpreventlist = ['*','?','|',':','<','>']
		sheetpreventresult = []
		for sheet_name_out1 in sheet_name_out1s:
			for sheetpreventlist1 in sheetpreventlist:
				if int(sheet_name_out1.find(sheetpreventlist1)) > -1:
					sheetpreventresult.append(str(sheet_name_out1.find(sheetpreventlist1)))	
		return sheetpreventresult
		
	def preventexcelcreate(file_string):
		sheet_name_out1s = list(file_string)
		sheetpreventlist = ['*','?','|','"','<','>']
		sheetpreventresult = []
		for sheet_name_out1 in sheet_name_out1s:
			for sheetpreventlist1 in sheetpreventlist:
				if int(sheet_name_out1.find(sheetpreventlist1)) > -1:
					sheetpreventresult.append(str(sheet_name_out1.find(sheetpreventlist1)))	
		return sheetpreventresult	

	if commandKey == "ReadFile":	
		try:
			folderpath =  sys.argv[2]
			def readfile(filename):
				text_file = open(filename,'r')
				#open the file
				#text_file = open('/Users/pankaj/abc.txt','r')

				#get the list of line
				line_list = text_file.readlines();
				text_file.close()
				text_array = []
				#for each line from the list, print the line
				for line in line_list:
					text_array.append(line)
				
				 #don't forget to close the file
				text_file1 = '\n'.join(str(e) for e in text_array)
				data = {}
				data['data'] = text_file1
				data['status_code'] = "200 OK"
				#data['status'] = "File Read successfully"
				data = json.dumps(data)
				return data	
			read_content = readfile(folderpath)
			read_content1 = ''.join(str(e) for e in read_content)
			print(read_content1)
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass	
		
	elif commandKey == "GetFolderContent":
		try:
			if os.path.exists(sys.argv[2]): 
				folderpath =  sys.argv[2]
				def list_file_and_folders(folderpath):
					file_array = []
					folder_array = []
					for root, dirs, files in os.walk(folderpath, topdown=False):
						for name in files:
							file_array.append(os.path.join(root, name))

						for name in dirs:
							#print("this is folder")
							folder_array.append(os.path.join(root, name))
							#stuff
					file_str = "Files"
					folder_str = "Folders"
					file_list = ','.join(str(e) for e in file_array)	
					folder_list = ','.join(str(e) for e in folder_array)	
					data = {}
					data['data'] = [file_str, file_list, folder_str, folder_list]
					data['status_code'] = "200 OK"
					#data['status'] = "File Content successfully"
					data = json.dumps(data)
					return data	
				list_file_folders = list_file_and_folders(folderpath)	
				list_file_folder = ''.join(str(e) for e in list_file_folders)
				print(list_file_folder)
				
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper folder path"
				print(json.dumps(data))
				pass
			
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		

	elif commandKey == "CopyFile":	
		try:
			folder_path_in =  sys.argv[2]
			folder_path_out =  sys.argv[3]	
			folder_path_out = folder_path_out.strip()
			head, tail = ntpath.split(folder_path_out)
			tail = tail.strip()
			folder_path_out = os.path.join(head,tail)			
			folder_path_out1 = folder_path_out			
			if os.path.exists(folder_path_in): 	
				if os.path.isfile(folder_path_in) or os.path.isfile(folder_path_out):
					filename_in1, file_extension_in1 =  sys.argv[2].split(".")
					filename_out, file_extension_out = sys.argv[3].split(".")
					if file_extension_in1 == file_extension_out :
						if os.path.exists(folder_path_out1): 
							data = {}
							data['status_code'] = "401"
							data['status'] = "Folder name already exists"
							print(data)
							pass
								
						else: 					
							def copy_paste_file(folder_path_in, folder_path_out1):
								#shutil.move(folder_path_in, folder_path_out1)		
								#os.rename(folder_path_in, folder_path_out1)
								shutil.copy2(folder_path_in, folder_path_out1)
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "File Coppied successfully"
								data = json.dumps(data)
								return data	
								
							file_move_txt = copy_paste_file(folder_path_in, folder_path_out1)	
							print(file_move_txt)
							
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass			
						
							
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide File formate only"
					print(json.dumps(data))
					pass	
					
			else:		
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide existing file"
				print(json.dumps(data))
				pass
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
		
	elif commandKey == "DeleteFile":
		try:
			argv2 =  sys.argv[2]	
			def delete_file(argv2):
				#shutil.os.remove(argv2)	
				os.remove(argv2)
				data = {}
				data['status_code'] = "200 OK"
				data['status'] = "File Deleted successfully"
				data = json.dumps(data)
				return data	
				
			file_delete_txt = delete_file(argv2)	
			print(file_delete_txt)	
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
		
	elif commandKey == "CreateFile":	
		try:
			argv2 =  sys.argv[2]
			argv2 = argv2.strip()
			head, tail = ntpath.split(argv2)
			tail = tail.strip()
			argv2 = os.path.join(head,tail)
			if os.path.exists(argv2): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				validatesheetname = preventexcelcreate(argv2)
				if len(validatesheetname) < 1:
					filename_in1, file_extension_in1 =  argv2.split(".")
					if file_extension_in1 not in (""):
						if filename_in1[-1] not in ('','/','\\',':'):
							if filename_in1[-1] not in (' '):			
								filename, file_extension = os.path.splitext(argv2)
								if filename[-1] not in (' '):							
									def create_file(argv2):
										with open(argv2, "w") as f:
											f.write("")
											
										data = {}
										data['status_code'] = "200 OK"
										data['status'] = "File created successfully"
										data = json.dumps(data)
										return data
										
									
									open_file_data = create_file(argv2)
									print(open_file_data)	
										
								else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file name"
										print(json.dumps(data))
										pass	
										
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file name"
								print(json.dumps(data))
								pass
																			
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper file name"
							print(json.dumps(data))
							pass
									
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass
						
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Not able create file name with '*','<','>','?'"
					print(json.dumps(data))
					pass		
					
					
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass	
		
	elif commandKey == "BrowseFile":	
		try:
			def browse_file():	
				Tkinter.Tk().withdraw() # Close the root window
				in_path = tkFileDialog.askopenfilename()
				#print in_path
				if in_path != "":
					data = {}
					data['file'] = in_path
					#data['status'] = "File created successfully"
					data = json.dumps(data)
					print(data)
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please select a file"
					print(json.dumps(data))
					pass		
					
			browse_file_path = browse_file()

			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
						
			
	elif commandKey == "CreateFolder":
		try:
			argv2 =  sys.argv[2]	
			validatesheetname = preventexcelcreate(sys.argv[2])
			if len(validatesheetname) < 1:
				if "/" in argv2:
					folder_path_splits = argv2.split("/")
									
					folder_path_splits1 = []	
					for folder_path_split in folder_path_splits:
						folder_path_splits1.append(folder_path_split.strip())					
					folder_path = '/'.join(str(e) for e in folder_path_splits1)	
					
				elif "\\" in argv2:
					folder_path_splits = argv2.split("\\")								
					folder_path_splits1 = []	
					for folder_path_split in folder_path_splits:
						folder_path_splits1.append(folder_path_split.strip())					
					folder_path = '/'.join(str(e) for e in folder_path_splits1)	
					
				else:
					folder_path = argv2.strip()
					
				
				if os.path.exists(folder_path): 
					data = {}
					data['status_code'] = "401"
					data['status'] = "Folder name already exists"
					print(json.dumps(data))
					pass
						
				else: 					
					def create_folder(folder_path):
						os.makedirs(folder_path)
						
						data = {}
						data['status_code'] = "200 OK"
						data['status'] = "Folder Created successfully"
						data = json.dumps(data)
						return data	
						
					folder_output_txt = create_folder(folder_path)	
					print(folder_output_txt)
					
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Not able create folder name with '*','<','>','?'"
				print(json.dumps(data))
				pass	
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
		
		
		
	elif commandKey == "CopyFolder":
		try:
		
			folder_path_in =  sys.argv[2]
			folder_path_out =  sys.argv[3]
			
			if "/" in folder_path_out:
				folder_path_splits = folder_path_out.split("/")
								
				folder_path_splits1 = []	
				for folder_path_split in folder_path_splits:
					folder_path_splits1.append(folder_path_split.strip())					
				folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
				
			elif "\\" in folder_path_out:
				folder_path_splits = folder_path_out.split("\\")								
				folder_path_splits1 = []	
				for folder_path_split in folder_path_splits:
					folder_path_splits1.append(folder_path_split.strip())					
				folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
				
			else:
				folder_path_out1 = folder_path_out.strip()
	
			
			if os.path.exists(folder_path_out1): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "Folder name already exists"
				print(json.dumps(data))
				pass
					
			else: 					
				def copy_folder(folder_path_in, folder_path_out1):
					shutil.copytree(folder_path_in, folder_path_out1)
					
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "Folder Copied successfully"
					data = json.dumps(data)
					return data	
					
				folder_copy_txt = copy_folder(folder_path_in, folder_path_out1)	
				print(folder_copy_txt)
				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass	

			
		
	elif commandKey == "MoveFolder":
		try:
		
			folder_path_in =  sys.argv[2]
			folder_path_out =  sys.argv[3]			

			if os.path.isfile(folder_path_in) or os.path.isfile(folder_path_out):
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide Folder formate only"
				print(data)
				pass	
			else:
				if "/" in folder_path_out:
					folder_path_splits = folder_path_out.split("/")
									
					folder_path_splits1 = []	
					for folder_path_split in folder_path_splits:
						folder_path_splits1.append(folder_path_split.strip())					
					folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
					
				elif "\\" in folder_path_out:
					folder_path_splits = folder_path_out.split("\\")								
					folder_path_splits1 = []	
					for folder_path_split in folder_path_splits:
						folder_path_splits1.append(folder_path_split.strip())					
					folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
					
				else:
					folder_path_out1 = folder_path_out.strip()
		
				
				if os.path.exists(folder_path_out1): 
					data = {}
					data['status_code'] = "401"
					data['status'] = "Folder name already exists"
					print(json.dumps(data))
					pass
						
				else: 					
					def move_folder(folder_path_in, folder_path_out1):
						#shutil.move(folder_path_in, folder_path_out1)		
						os.rename(folder_path_in, folder_path_out1)
						data = {}
						data['status_code'] = "200 OK"
						data['status'] = "Folder moved successfully"
						data = json.dumps(data)
						return data	
						
					folder_move_txt = move_folder(folder_path_in, folder_path_out1)	
					print(folder_move_txt)
						
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
				
				
				
	elif commandKey == "MoveFile":
		try:
		
			folder_path_in =  sys.argv[2]
			folder_path_out =  sys.argv[3]	
			folder_path_out = folder_path_out.strip()
			head, tail = ntpath.split(folder_path_out)
			tail = tail.strip()
			folder_path_out = os.path.join(head,tail)			
			folder_path_out1 = folder_path_out			
			if os.path.exists(folder_path_in): 	
				if os.path.isfile(folder_path_in) or os.path.isfile(folder_path_out):
					filename_in1, file_extension_in1 =  sys.argv[2].split(".")
					filename_out, file_extension_out = sys.argv[3].split(".")
					if file_extension_in1 == file_extension_out :
						if os.path.exists(folder_path_out1): 
							data = {}
							data['status_code'] = "401"
							data['status'] = "Folder name already exists"
							print(data)
							pass
								
						else: 					
							def move_file(folder_path_in, folder_path_out1):
								#shutil.move(folder_path_in, folder_path_out1)		
								os.rename(folder_path_in, folder_path_out1)
								data = {}
								data['status_code'] = "200 OK"
								data['status'] = "File Moved successfully"
								data = json.dumps(data)
								return data	
								
							file_move_txt = move_file(folder_path_in, folder_path_out1)	
							print(file_move_txt)
							
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass			
						
							
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide File formate only"
					print(json.dumps(data))
					pass	
					
			else:		
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide existing file"
				print(json.dumps(data))
				pass		

		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
							
			
			
	elif commandKey == "WriteFile":	
		try:
			if os.path.exists(sys.argv[2]): 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				filename_in1, file_extension_in1 =  sys.argv[2].split(".")
				if filename_in1 not in (""):
					if filename_in1[-1] not in ('','/','\\'):		
						if filename_in1[-1] not in (' '):			
							if filename_in[-1] not in (' '):	
								if sys.argv[2].endswith('txt'):							
									argv2 =  sys.argv[2]
									write_doc_txt =  sys.argv[3]
									def write_file(argv2):
										with open(argv2, "w") as f:
											f.write(write_doc_txt)
											
										data = {}
										data['status_code'] = "200 OK"
										data['status'] = "File written successfully"
										data = json.dumps(data)
										return data
																		
									write_file_data = write_file(argv2)
									print(write_file_data)		
											
								else:	
									data = {}
									data['status_code'] = "401"
									data['status'] = "Supporting only .txt formate for write"
									data = json.dumps(data)
									print(data)
									pass	
	
								
							else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please provide proper file name"
									print(json.dumps(data))
									pass	
								
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper file name"
							print(json.dumps(data))
							pass
																		
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass
								
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide proper file extension"
					print(json.dumps(data))
					pass
					
					
			else:		
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide existing file"
				print(json.dumps(data))
				pass
					
		 		
					
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
			
			
	elif commandKey == "DeleteFolder":
		try:
		
			#folder_path_in =  sys.argv[2]
			folder_path_out =  sys.argv[2]
			
			if "/" in folder_path_out:
				folder_path_splits = folder_path_out.split("/")
								
				folder_path_splits1 = []	
				for folder_path_split in folder_path_splits:
					folder_path_splits1.append(folder_path_split.strip())					
				folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
				
			elif "\\" in folder_path_out:
				folder_path_splits = folder_path_out.split("\\")								
				folder_path_splits1 = []	
				for folder_path_split in folder_path_splits:
					folder_path_splits1.append(folder_path_split.strip())					
				folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
				
			else:
				folder_path_out1 = folder_path_out.strip()
				
	
			if os.path.exists(folder_path_out1): 					
				def delete_folder(folder_path_out1):
					shutil.rmtree(folder_path_out1)
					#os.removedirs(folder_path_out1)
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "Folder Deleted successfully"
					data = json.dumps(data)
					return data	
					
				folder_delete_txt = delete_folder(folder_path_out1)	
				print(folder_delete_txt)
				
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide existing folder"
				print(json.dumps(data))
				pass			
				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "CheckExistenceFolder":
		try:
			folder_path_out =  sys.argv[2]
			
			if "/" in folder_path_out:
				folder_path_splits = folder_path_out.split("/")
								
				folder_path_splits1 = []	
				for folder_path_split in folder_path_splits:
					folder_path_splits1.append(folder_path_split.strip())					
				folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
				
			elif "\\" in folder_path_out:
				folder_path_splits = folder_path_out.split("\\")								
				folder_path_splits1 = []	
				for folder_path_split in folder_path_splits:
					folder_path_splits1.append(folder_path_split.strip())					
				folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
				
			else:
				folder_path_out1 = folder_path_out.strip()
				
			folder_path_out =  folder_path_out1
			
			if os.path.exists(folder_path_out): 					
				def existance_folder():					
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "Folder existed"
					data = json.dumps(data)
					return data	
					
				folder_existance_txt = existance_folder()	
				print(folder_existance_txt)
				
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Folder doesn't existed"
				print(json.dumps(data))
				pass			
				
				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
			
			
	elif commandKey == "CheckExistenceFile":
		try:
			folder_path_out =  sys.argv[2]
			filename_in1, file_extension_in1 =  sys.argv[2].split(".")
			if filename_in1 not in (""):
				if os.path.exists(folder_path_out): 					
					def existance_file():					
						data = {}
						data['status_code'] = "200 OK"
						data['status'] = "File existed"
						data = json.dumps(data)
						return data	
						
					file_existance_txt = existance_file()	
					print(file_existance_txt)
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "File doesn't existed"
					print(json.dumps(data))
					pass	

			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper file extension"
				print(json.dumps(data))
				pass					

				
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "ZipFolder":
		try:
			folder_path = sys.argv[2]
			folder_path_out =  sys.argv[2]
			zip_output_file =  sys.argv[3]
			if "/" in folder_path_out:
				folder_path_splits = folder_path_out.split("/")
								
				folder_path_splits1 = []	
				for folder_path_split in folder_path_splits:
					folder_path_splits1.append(folder_path_split.strip())					
				folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
				
			elif "\\" in folder_path_out:
				folder_path_splits = folder_path_out.split("\\")								
				folder_path_splits1 = []	
				for folder_path_split in folder_path_splits:
					folder_path_splits1.append(folder_path_split.strip())					
				folder_path_out1 = '/'.join(str(e) for e in folder_path_splits1)	
				
			else:
				folder_path_out1 = folder_path_out.strip()
				
			folder_path_out =  folder_path_out1
			if os.path.isfile(folder_path):
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper folder name"
				print(json.dumps(data))
				pass
			else:
				if os.path.exists(folder_path_out): 
					if sys.argv[3].endswith('.zip') or sys.argv[3].endswith('.rar'):
						zf = zipfile.ZipFile(zip_output_file, "w")
						for dirname, subdirs, files in os.walk(folder_path):
							zf.write(dirname)
							for filename in files:
								zf.write(os.path.join(dirname, filename))
						zf.close()	
						data = {}
						data['status_code'] = "200 OK"
						data['status'] = "Folder zipped successfully"
						data = json.dumps(data)
						print(data)
						
					elif zip_output_file in ('','/'):
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file name"
						print(json.dumps(data))
						pass	
						
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper archive file extension"
						print(json.dumps(data))
						pass
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Folder doesn't existed"
					print(json.dumps(data))
					pass
				
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "ZipFile":
		try:
			folder_path =  sys.argv[2]
			zip_output_file =  sys.argv[3]
			if os.path.isfile(folder_path):
				if os.path.exists(sys.argv[2]): 
					filename_in, file_extension_in = os.path.splitext(sys.argv[2])
					if file_extension_in not in (""):
						filename_out, file_extension_out = sys.argv[2].split(".")
						if filename_out not in (""):
							if filename_out[-1] not in ('','/'):
								if filename_out[-1] not in (' '):
									if sys.argv[3].endswith('.zip') or sys.argv[3].endswith('.rar'):
										if folder_path[-1] not in (' '):
											with ZipFile(zip_output_file, 'w') as myzip:
												myzip.write(folder_path)
											myzip.close()	
											data = {}
											data['status_code'] = "200 OK"
											data['status'] = "File zipped successfully"
											data = json.dumps(data)
											print(data)
										else:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper file name"
											print(json.dumps(data))
											pass	
												
									elif filename_out[-1] in ('','/'):
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file name"
										print(json.dumps(data))
										pass	
										
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper archive file extension"
										print(json.dumps(data))
										pass
										
								else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file name"
										print(json.dumps(data))
										pass	
										
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file name"
								print(json.dumps(data))
								pass		
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper file name"
							print(json.dumps(data))
							pass
							
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass
						
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide Existing file name"
					print(json.dumps(data))
					pass
												
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper file extension"
				print(json.dumps(data))
				pass
				

		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
					
			
	elif commandKey == "PrintFile":
		try:
			file_path =  sys.argv[2]
			if os.path.isfile(file_path):
				if os.path.exists(file_path): 
					os.startfile(file_path, "print")	
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "Task completed successfully"
					data = json.dumps(data)
					print(data)
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide Existing file name"
					print(json.dumps(data))
					pass
												
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper file extension"
				print(json.dumps(data))
				pass
				

		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass	
			
			
	elif commandKey == "PrintMultipleFiles":
		try:
			folder_path =  sys.argv[2]
			if os.path.isfile(folder_path):
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper folder path"
				print(json.dumps(data))
				pass
				
			else:
				if os.path.exists(folder_path): 
					files = os.listdir(folder_path)
					os.chdir(folder_path)
					for file in glob.glob("*.*" ):
						#print(file)
						if (file != ''):
							os.startfile(file, "print")
							time.sleep(6)
								
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "Task completed successfully"
					data = json.dumps(data)
					print(data)	
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide existing folder path"
					print(json.dumps(data))
					pass

					
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
			
			
	elif commandKey == "PrintMultipleFilesWithExtension":
		try:
			folder_path =  sys.argv[2]
			extention_name =  sys.argv[3]
			if os.path.isfile(folder_path):
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please provide proper folder path"
				print(json.dumps(data))
				pass
				
			else:
				if os.path.exists(folder_path): 
					files = os.listdir(folder_path)
					'''
					for file in files:
						os.startfile(file, "print")	
						time.sleep(2)		
					'''
					os.chdir(folder_path)
					filename_in, file_extension_in = os.path.splitext(sys.argv[2])
					
					for file in glob.glob("*.*" ):
						#print(file)
						if (file != ''):
							filename_in, file_extension_in = os.path.splitext(file)
							#print(file_extension_in)
							if file_extension_in in (extention_name):
								os.startfile(file, "print")
								time.sleep(6)
								
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "Task completed successfully"
					data = json.dumps(data)
					print(data)	
					
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide existing folder path"
					print(json.dumps(data))
					pass

					
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
			
	elif commandKey == "GetPrintersList":
		try:
			printer_info = win32print.EnumPrinters (
				win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
			)
			printer_names = [name for (flags, description, name, comment) in 
			printer_info]
			print_list = []
			for i, name in enumerate (printer_names):
			   print_list.append((name))
			data = {}
			data['status_code'] = "200 OK"
			data['data'] = print_list
			data = json.dumps(data)
			print(data)	
					
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
					
	elif commandKey == "RenameFile":
		try:
		    
			folder_path_in =  sys.argv[2]
			inputFilepath = sys.argv[3]
			filename_w_ext = os.path.basename(inputFilepath)
			filename, file_extension = os.path.splitext(filename_w_ext)
			path, filename = os.path.split(inputFilepath)
			validatesheetname = preventexcelrename(filename)
			if len(validatesheetname) < 1:
				folder_path_in = folder_path_in.strip()
				head1, tail1 = ntpath.split(folder_path_in)
				tail1 = tail1.strip()
				folder_path_out =  sys.argv[3]	
				folder_path_out = folder_path_out.strip()
				head, tail = ntpath.split(folder_path_out)
				tail = tail.strip()
				folder_path_out = os.path.join(head,tail)			
				folder_path_out1 = folder_path_out			
				if os.path.exists(folder_path_in): 	
					if head1 == head :
						if os.path.isfile(folder_path_in) or os.path.isfile(folder_path_out):
							filename_in1, file_extension_in1 =  sys.argv[2].split(".")
							filename_out, file_extension_out = sys.argv[3].split(".")
							if file_extension_in1 == file_extension_out :
								if os.path.exists(folder_path_out1): 
									data = {}
									data['status_code'] = "401"
									data['status'] = "Folder name already exists"
									print(data)
									pass
										
								else: 					
									def move_file(folder_path_in, folder_path_out1):
										shutil.move(folder_path_in, folder_path_out1)					
										data = {}
										data['status_code'] = "200 OK"
										data['status'] = "File renamed successfully"
										data = json.dumps(data)
										return data	
										
									file_move_txt = move_file(folder_path_in, folder_path_out1)	
									print(file_move_txt)
									
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file extension"
								print(json.dumps(data))
								pass			
								
									
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide File formate only"
							print(json.dumps(data))
							pass	
							
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Rename file location not match"
						print(json.dumps(data))
						pass	
						
				else:		
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide existing file"
					print(json.dumps(data))
					pass		

			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Not able rename file name with '<','>','*',,':','?'"
				print(json.dumps(data))
				pass		
			
			
					
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass	
			
			
	elif commandKey == "UnzipFile":
		try:
			folder_path =  sys.argv[2]
			zip_output_file =  sys.argv[3]
			zip_output_file = zip_output_file.strip()
			head, tail = ntpath.split(zip_output_file)
			tail = tail.strip()
			zip_output_file = os.path.join(head,tail)
			if os.path.exists(zip_output_file): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
			else:
				if os.path.exists(folder_path): 
					filename_in, file_extension_in = os.path.splitext(sys.argv[2])
					if file_extension_in not in (""):
						filename_out, file_extension_out = sys.argv[2].split(".")
						if filename_out not in (""):
							if filename_out[-1] not in ('','/'):
								if filename_out[-1] not in (' '):
									if sys.argv[2].endswith('.zip') or sys.argv[2].endswith('.rar'):
										if folder_path[-1] not in (' '):
											zip_ref = zipfile.ZipFile(folder_path, 'r')
											zip_ref.extractall(zip_output_file)
											zip_ref.close()	
											data = {}
											data['status_code'] = "200 OK"
											data['status'] = "File unzipped successfully"
											data = json.dumps(data)
											print(data)
										else:
											data = {}
											data['status_code'] = "401"
											data['status'] = "Please provide proper file name"
											print(json.dumps(data))
											pass	
									
									elif filename_out[-1] in ('','/'):
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file name"
										print(json.dumps(data))
										pass	
										
									else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper archive file extension"
										print(json.dumps(data))
										pass
										
								else:
										data = {}
										data['status_code'] = "401"
										data['status'] = "Please provide proper file name"
										print(json.dumps(data))
										pass	
										
							else:
								data = {}
								data['status_code'] = "401"
								data['status'] = "Please provide proper file name"
								print(json.dumps(data))
								pass		
							
						else:
							data = {}
							data['status_code'] = "401"
							data['status'] = "Please provide proper file name"
							print(json.dumps(data))
							pass
							
					else:
						data = {}
						data['status_code'] = "401"
						data['status'] = "Please provide proper file extension"
						print(json.dumps(data))
						pass
						
				else:
					data = {}
					data['status_code'] = "401"
					data['status'] = "Please provide Existing file name"
					print(json.dumps(data))
					pass


		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass		
					
					
			
	else:
		print("Please enter proper Command Key")	
		
		
except Exception as e:
	data = {}
	data['status_code'] = "401"
	data['status'] = str(e) 
	print(json.dumps(data))
	pass
	#print("please enter proper arguments")	