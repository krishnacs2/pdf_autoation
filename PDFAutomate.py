import PyPDF2
import sys
import json
import shutil
import os.path
import requests
import ntpath

try:
	# creating an object 
	commandKey = sys.argv[1]
	
	
	#OCR for select text or image in PDF
	commandKey = sys.argv[1]
	payload = {'isOverlayRequired': "false",
				'apikey': "61d20edaae88957",
				'language': "eng",
				}
	
	def readOcrText(file_name):
		dataArray = [] 
		filename = file_name
		with open(filename, 'rb') as f:
		   r = requests.post('https://api.ocr.space/parse/image',files={filename: f},data=payload) 
		#print(r.content.decode())
		result = json.loads(r.content.decode('utf-8'))
		#print(result['ParsedResults'])			
		row_json = json.dumps(result['ParsedResults'][0]['ParsedText'])
		#print(row_json)	
		row_json_result = str(row_json)
		row_json_result = row_json_result.replace('"',"")
		row_json_result = row_json_result
		#print(row_json_result)
		#dataArray.append(image+"   "+response.json()["best_label"]["label_name"]+":"+row_json_result)	  
		#print(response.json()["best_label"]["label_name"]+":"+row_json_result)
		mylist = row_json_result.split('\\r\\n')
		str1 = '\n '.join(str(e) for e in mylist)
		data = {}
		data['status_code'] = "200 OK"
		data['data'] = str1
		ocr_txt = json.dumps(data)		
			
		return ocr_txt
		
		
	
	#file = open('text2.pdf', 'rb')
	if commandKey == "ReadPdfbyPage":	
		try:
			file_name =  sys.argv[2] 
			page_number =  int(sys.argv[3]) - 1 
			if page_number + 1  > 0:
				def ReadPdfbyPage(file_name):
					# creating a pdf reader object
					file = open(file_name, 'rb')
					fileReader = PyPDF2.PdfFileReader(file)

					pageObj = fileReader.getPage(page_number)
					return pageObj.extractText().encode('utf-8').strip()
					

				
				pdf_text = ReadPdfbyPage(file_name)
				data = {}
				data['status_code'] = "200 OK"
				data['data'] = str(pdf_text)
				print(json.dumps(data))
			
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please enter proper page number"
				print(json.dumps(data))
				pass			
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
	elif commandKey == "WritePdfbyPage":	
		try:
			if os.path.exists(sys.argv[4]): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				#filename_out, file_extension_out = os.path.splitext(sys.argv[3])
				filename_out, file_extension_out = sys.argv[4].split(".")
				#print(filename_out[-1])
				if filename_out not in (""):
					if filename_out[-1] not in ('','/'):
						if filename_out[-1] not in (' '):
							if sys.argv[2].endswith('.pdf') and sys.argv[4].endswith('.pdf'):
								file_name =  sys.argv[2] 
								page_number =  int(sys.argv[3]) - 1
								if page_number + 1  > 0:
									new_file_name = sys.argv[4]
									
									def WritePdfbyPage(file_name):
										# creating a pdf reader object
										pdfFile = open(file_name, 'rb')
										pdfReader = PyPDF2.PdfFileReader(pdfFile)
										pdfWriter = PyPDF2.PdfFileWriter()
										pageObj = pdfReader.getPage(page_number)
										pdfWriter.addPage(pageObj)
										pdfOutputFile = open(new_file_name, 'wb')
										pdfWriter.write(pdfOutputFile)
										pdfOutputFile.close()
										pdfFile.close()
										data = {}
										data['status_code'] = "200 OK"
										data['status'] = "File Written successfully"
										data = json.dumps(data)
										return data
									
									pdf_conttent = WritePdfbyPage(file_name)
									print(pdf_conttent)
									
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please enter proper page number"
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
								data['status'] = "Please provide proper file extension"
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
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
	elif commandKey == "ReadPdfPagebyRange":	
		try:
			file_name =  sys.argv[2] 
			temp_list1 = sys.argv[3]
			temp_list2 = temp_list1.split(',')
			if int(temp_list2[0])  > 0 and int(temp_list2[1]) > 0 and int(temp_list2[1]) > int(temp_list2[0]):
				def ReadPdfPagebyRange(file_name):
					# creating a pdf reader object
					pdfFile = open(file_name, 'rb')
					pdfReader = PyPDF2.PdfFileReader(pdfFile)
					text = ""
					for pageNum in range(int(temp_list2[0])-1,int(temp_list2[1])):
						 pageObj = pdfReader.getPage(pageNum)
						 text += pageObj.extractText()
						 text += "\n"
					return text

				pdf_text = ReadPdfPagebyRange(file_name)
				data = {}
				data['status_code'] = "200 OK"
				data['data'] = str(pdf_text)
				print(json.dumps(data))
						    						
			else:
				data = {}
				data['status_code'] = "401"
				data['status'] = "Please enter proper page range"
				print(json.dumps(data))
				pass	
				

		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass	
			
	elif commandKey == "WritePdfPagebyRange":	
		try:
			if os.path.exists(sys.argv[4]): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				#filename_out, file_extension_out = os.path.splitext(sys.argv[3])
				filename_out, file_extension_out = sys.argv[4].split(".")
				#print(filename_out[-1])
				if filename_out not in (""):
					if filename_out[-1] not in ('','/'):
						if filename_out[-1] not in (' '):
							if sys.argv[2].endswith('.pdf') and sys.argv[4].endswith('.pdf'):
								file_name =  sys.argv[2] 
								temp_list1 = sys.argv[3]
								new_file_name = sys.argv[4]
								temp_list2 = temp_list1.split(',')
								if int(temp_list2[0])  > 0 and int(temp_list2[1]) > 0 and int(temp_list2[1]) > int(temp_list2[0]):
									def WritePdfPagebyRange(file_name):
										# creating a pdf reader object
										pdfFile = open(file_name, 'rb')
										pdfReader = PyPDF2.PdfFileReader(pdfFile)
										pdfWriter = PyPDF2.PdfFileWriter()
										for pageNum in range(int(temp_list2[0])-1,int(temp_list2[1])):
											pageObj = pdfReader.getPage(pageNum)
											pdfWriter.addPage(pageObj)
											
										pdfOutputFile = open(new_file_name, 'wb')
										pdfWriter.write(pdfOutputFile)
										pdfOutputFile.close()
										pdfFile.close()
										data = {}
										data['status_code'] = "200 OK"
										data['status'] = "File Written successfully"
										data = json.dumps(data)
										return data
									
									pdf_text = WritePdfPagebyRange(file_name)
									print(pdf_text)
									
								else:
									data = {}
									data['status_code'] = "401"
									data['status'] = "Please enter proper page range"
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
								data['status'] = "Please provide proper file extension"
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
								
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
	elif commandKey == "ReadPdf":	
		try:
			file_name =  sys.argv[2] 
			
			def ReadPdf(file_name):
				# creating a pdf reader object
				pdfFile = open(file_name, 'rb')
				pdfReader = PyPDF2.PdfFileReader(pdfFile)
				text = ""
				for pageNum in range(pdfReader.numPages):
					 pageObj = pdfReader.getPage(pageNum)
					 text += pageObj.extractText()
					 text += "\n"
				return text
			
			pdf_text = ReadPdf(file_name)
			data = {}
			data['status_code'] = "200 OK"
			data['data'] = str(pdf_text)
			print(json.dumps(data))
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass	
			
	elif commandKey == "GetPdfPageCount":	
		try:
			file_name =  sys.argv[2] 
			
			def GetPdfPageCount(file_name):
				# creating a pdf reader object
				file = open(file_name, 'rb')
				fileReader = PyPDF2.PdfFileReader(file)

				# print the number of pages in pdf file
				page_count = fileReader.numPages
				#text = textract.process('text2.pdf', method='tesseract',language='eng')
				#print(text)
				data = {}
				data['Total Pages'] = page_count
				data = json.dumps(data)
				return data
				return pageObj.extractText()
			
			pdf_number = GetPdfPageCount(file_name)
			data = {}
			data['status_code'] = "200 OK"
			data['data'] = str(pdf_number)
			print(json.dumps(data))
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
	elif commandKey == "CopyPdf":	
		try:
			output_file_name =  sys.argv[3] 
			output_file_name = output_file_name.strip()
			head, tail = ntpath.split(output_file_name)
			tail = tail.strip()
			output_file_name = os.path.join(head,tail)
			if os.path.exists(output_file_name): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				#filename_out, file_extension_out = os.path.splitext(sys.argv[3])
				filename_out, file_extension_out = sys.argv[3].split(".")
				#print(filename_out[0])
				if filename_out not in (""):
					if filename_out[-1] not in ('','/'):
						if filename_out[-1] not in (' '):
							if sys.argv[2].endswith('.pdf') and sys.argv[3].endswith('.pdf'):
								file_name =  sys.argv[2] 
								def copy_paste_file(file_name, output_file_name):
									shutil.copy2(file_name, output_file_name)
									#another way to copy file
									#shutil.copyfile('/Users/pankaj/abc.txt', '/Users/pankaj/abc_copyfile.txt')
									data = {}
									data['status_code'] = "200 OK"
									data['status'] = "PDF Copied successfully"
									data = json.dumps(data)
									return data
									
								copy_file_content = copy_paste_file(file_name, output_file_name)	
								print(copy_file_content)
								
							elif filename_out[-1] in ('','/'):
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
					
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass	
			
	elif commandKey == "MovePdf":	
		try:
			output_file_name =  sys.argv[3] 
			output_file_name = output_file_name.strip()
			head, tail = ntpath.split(output_file_name)
			tail = tail.strip()
			output_file_name = os.path.join(head,tail)
			if os.path.exists(output_file_name): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				#filename_out, file_extension_out = os.path.splitext(sys.argv[3])
				filename_out, file_extension_out = sys.argv[3].split(".")
				#print(filename_out[-1])
				if filename_out not in (""):
					if filename_out[-1] not in ('','/'):
						if filename_out[-1] not in (' '):
							if sys.argv[2].endswith('.pdf') and sys.argv[3].endswith('.pdf'):
								file_name =  sys.argv[2] 
								#output_file_name =  sys.argv[3] 
								def move_pdf_file(file_name, output_file_name):
									#shutil.move(file_name, output_file_name)
									os.rename(file_name, output_file_name)
									#another way to copy file
									#shutil.copyfile('/Users/pankaj/abc.txt', '/Users/pankaj/abc_copyfile.txt')
									data = {}
									data['status_code'] = "200 OK"
									data['status'] = "PDF Moved successfully"
									data = json.dumps(data)
									return data
									
								copy_file_content = move_pdf_file(file_name, output_file_name)	
								print(copy_file_content)
															
															
							elif filename_out[-1] in ('','/'):
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
			
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass				
			
	elif commandKey == "DeletePdf":
		try:
			if sys.argv[2].endswith('.pdf'): 
				argv2 =  sys.argv[2]	
				argv2 = argv2.strip()
				head, tail = ntpath.split(argv2)
				tail = tail.strip()
				argv2 = os.path.join(head,tail)
				def delete_file(argv2):
					shutil.os.remove(argv2)	
					
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "File Deleted successfully"
					data = json.dumps(data)
					return data	
					
				file_delete_txt = delete_file(argv2)	
				print(file_delete_txt)	
						
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
			
	elif commandKey == "MergePdffiles":	
		try:
			output_file_name = sys.argv[4]
			output_file_name = output_file_name.strip()
			head, tail = ntpath.split(output_file_name)
			tail = tail.strip()
			output_file_name = os.path.join(head,tail)
			if os.path.exists(output_file_name): 
				data = {}
				data['status_code'] = "401"
				data['status'] = "File name already exists"
				print(json.dumps(data))
				pass
					
			else: 
				filename_in, file_extension_in = os.path.splitext(sys.argv[2])
				filename_in1, file_extension_in1 = os.path.splitext(sys.argv[3])
				#filename_out, file_extension_out = os.path.splitext(sys.argv[3])
				filename_out, file_extension_out = sys.argv[4].split(".")
				#print(filename_out[-1])
				if filename_out not in (""):
					if filename_out[-1] not in ('','/'):
						if filename_out[-1] not in (' '):
							if sys.argv[2].endswith('.pdf') and sys.argv[3].endswith('.pdf') and sys.argv[4].endswith('.pdf'):
								file_name1 =  sys.argv[2] 
								file_name2 = sys.argv[3]
								

								def mergePdffiles(file_name1, file_name2):
									# creating a pdf reader object
									pdf1File = open(file_name1, 'rb')
									pdf2File = open(file_name2, 'rb')
									pdf1Reader = PyPDF2.PdfFileReader(pdf1File)
									pdf2Reader = PyPDF2.PdfFileReader(pdf2File)
									pdfWriter = PyPDF2.PdfFileWriter()
									
									for pageNum in range(pdf1Reader.numPages):
										 pageObj = pdf1Reader.getPage(pageNum)
										 pdfWriter.addPage(pageObj)

									for pageNum in range(pdf2Reader.numPages):
										 pageObj = pdf2Reader.getPage(pageNum)
										 pdfWriter.addPage(pageObj)
										
									pdfOutputFile = open(output_file_name, 'wb')
									pdfWriter.write(pdfOutputFile)
									pdfOutputFile.close()
									pdf1File.close()
									pdf2File.close()
									data = {}
									data['status_code'] = "200 OK"
									data['status'] = "PDF Files Merged successfully"
									data = json.dumps(data)
									return data
								
								pdf_text = mergePdffiles(file_name1, file_name2)
								print(pdf_text)
								
							elif filename_out[-1] in ('','/'):
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
		
		except OSError as e:
			data = {}
			data['status_code'] = "401"
			data['status'] = str(e) 
			print(json.dumps(data))
			pass
			
	elif commandKey == "ClosePdf":
		try:
			if sys.argv[2].endswith('.pdf'): 
				file_name =  sys.argv[2]	
				file_name = file_name.strip()
				head, tail = ntpath.split(file_name)
				tail = tail.strip()
				file_name = os.path.join(head,tail)
				def closePdf(argv2):
					pdfFile = open(file_name, 'rb')
					pdfFile.close()
					data = {}
					data['status_code'] = "200 OK"
					data['status'] = "PDF File Closed successfully"
					data = json.dumps(data)
					return data	
					
				file_close_txt = closePdf(file_name)	
				print(file_close_txt)
				
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
			
	
	elif commandKey == "SelectArea":
		try:	
			file_name1 = sys.argv[2]	
			file_name1 = file_name1.strip()	
			head, tail = ntpath.split(file_name1)
			tail = tail.strip()
			file_name1 = os.path.join(head,tail)
			if os.path.exists(file_name1): 
				filename, file_extension = os.path.splitext(file_name1)
				if filename[-1] not in (' '):
					if file_extension in ('.PNG', '.png','.jpg','.JPG','.jpeg','.JPEG','.GIF','.gif'):		
						#file_name1 = sys.argv[2]	
						#filedir = "upload" 
						#filename = "process.png"
						#file_name = filename
						#shutil.copy2(file_name1, file_name)			
						ocr_content = readOcrText(sys.argv[2])	
						print(ocr_content)
					elif file_extension in (''):
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
					data['status'] = "Please provide proper file name"
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
	
			
	else:
		print("Please enter proper Command Key")	
	
	
except Exception as e:
	data = {}
	data['status_code'] = "401"
	data['status'] = str(e) 
	#data = json.dumps(data)
	print(json.dumps(data))
	pass	