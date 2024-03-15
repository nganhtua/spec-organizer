import os, sqlite3, datetime, hashlib, shutil, html, webbrowser
from stat import S_IREAD
import diff_match_patch as dmp
import docx
import win32com.client as win32

class db_handler:
	def __init__(self, db):
		con = sqlite3.connect(db)
		con.close()
		self.db = db

	def table_info(self, table):
		# Print table columns infos
		con = sqlite3.connect(self.db)
		cur = con.cursor()
		cur.execute(f'PRAGMA table_info({table})')
		for i in cur.fetchall():
			print(i)
		con.close()

	def update_table(self, table, *lines):
		''' Update table data or append new lines. '''
		con = sqlite3.connect(self.db)
		cur = con.cursor()

		# Retrive the column names from the table
		cur.execute(f'PRAGMA table_info({table})')
		cols = [column[1] for column in cur.fetchall()]

		# Retrieve the primary key column names
		table_pri = cur.execute(f"SELECT rowid, name \
			FROM pragma_table_info(\'{table}\') WHERE pk <> 0").fetchall()
		pri_rowids = [i[0]-1 for i in table_pri]
		pri_keys = [i[1] for i in table_pri]

		# Process line by line
		for line in lines:
			print('\n...Processing {}...'.format(line))
			if len(line) != len(cols):
				# Columns numbers don't match
				print('Columns didn\'t match. There are {} columns in table {}. Your input was {}.'
					   .format(len(cols), table, len(line)))
				raise Exception
			else:
				# Columns numbers match
				where_clause = ' AND '.join(f"{pri_key}=?" for pri_key in pri_keys)
				where_lookup = [line[i] for i in pri_rowids]
				sql_cmd = f'SELECT * FROM {table} WHERE {where_clause}'
				print(sql_cmd, where_lookup)
				current = cur.execute(sql_cmd, where_lookup).fetchall()
				if len(current) > 0:
					# The key exists
					sql_cmd = f'UPDATE {table} SET '
					update_cols = []
					update_values = []
					for col, value in zip(cols, line):
						if col not in pri_keys and value is not None:
							update_cols.append(f'{col}=?')
							update_values.append(value)
					sql_cmd += ', '.join(update_cols)
					sql_cmd += f' WHERE {where_clause}'
					print(sql_cmd, update_values + where_lookup)
					cur.execute(sql_cmd, update_values + where_lookup)
				else:
					# The key doesn't exist, insert a new row
					placeholders = ', '.join(['?' for _ in cols])
					sql_cmd = f"INSERT INTO {table} ({', '.join(cols)}) \
								VALUES ({placeholders})"
					print(sql_cmd, line)
					cur.execute(sql_cmd, line)
		con.commit()
		con.close()

	def update_key(self, table, old_key, new_key):
		''' Update primary key values of a record. '''
		if len(old_key) != len(new_key):
			print('Keys mismatch.')
			raise Exception
		con = sqlite3.connect(self.db)
		cur = con.cursor()
		cur.execute(f'SELECT name FROM pragma_table_info("{table}")')
		cols = [i[0] for i in cur.fetchall()]
		cur.execute(f'SELECT name FROM pragma_table_info("{table}") \
					WHERE pk <> 0')
		pri_keys = [i[0] for i in cur.fetchall()]
		no_of_pri_cols = len(pri_keys)
		if len(old_key) != len(pri_keys):
			print('Filter values and primary key mismatch.')
		set_clause = ', '.join(f"{pri_key}=?" for pri_key in pri_keys)
		# set_values = new_key already
		where_clause = ' AND '.join(f"{pri_key}=?" for pri_key in pri_keys)
		# where_lookup = old_key already
		sql_cmd = f'UPDATE {table} SET {set_clause} WHERE {where_clause}'
		print(sql_cmd, new_key+old_key)
		cur.execute(sql_cmd, new_key+old_key)
		con.commit()
		con.close()

	# TABLE spec_list
	def process_file(self, path):
		''' Copy file to contents folder, name the copied file as MD5 value. 
		Return the MD5 of the file. '''
		try:
			with open(path, 'rb') as file:
				md5_hash = hashlib.md5(file.read()).hexdigest()
			newfile = md5_hash + os.path.splitext(path)[1]
			if os.path.isfile('contents/' + newfile):
				return newfile
			shutil.copy2(path, 'contents/' + newfile)
			os.chmod('contents/' + newfile, S_IREAD)
			return newfile
		except FileNotFoundError:
			print('FileNotFoundError')
			return None
		except Exception as e:
			print(e)
			return None

	def convert_doc_to_txt(self, doc_file, txt_file):
		''' Convert .doc or .docx file to .txt file, using UTF-8 encoding '''
		try:
			# Check if the file is a .docx document
			if doc_file.endswith('.docx'):
				doc = docx.Document(doc_file)
				text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
				with open(txt_file, 'w', encoding='utf-8') as file:
					file.write(text)
			else:
				# For .doc files, open and save it as txt file
				word_app = win32.gencache.EnsureDispatch('Word.Application')
				word_app.Visible = False
				full_path = os.path.abspath(doc_file)
				doc = word_app.Documents.Open(full_path)
				print(os.path.dirname(full_path))
				doc.SaveAs('/'.join([os.getcwd(), txt_file]), \
					FileFormat=win32.constants.wdFormatText, Encoding=65001)
				doc.Close()
				word_app.Quit()
			print(f'Conversion successful: {doc_file} -> {txt_file}')
		except Exception as e:
			print(f'Error occurred during conversion: {e}')

	def attach_to_spec(self, path, key, field):
		''' Attach a file to a record in spec_list table. '''
		allowed_field = {'pdf_file': ['.pdf'], 'doc_file':['.doc', '.docx'], \
						 'txt_file': ['.txt']}
		if field not in allowed_field.keys():
			print('Field {} doesn\'t exist in table spec_list.'.format(field))
			raise Exception
		con = sqlite3.connect(self.db)
		cur = con.cursor()
		cur.execute(f'SELECT name FROM pragma_table_info("spec_list") \
					WHERE pk <> 0')
		pri_keys = [i[0] for i in cur.fetchall()]
		no_of_pri_cols = len(pri_keys)

		if len(key) != len(pri_keys):
			print('Key length mismatch.')
			raise Exception
		where_clause = ' AND '.join([f'{pri_key}=?' for pri_key in pri_keys])\
		# where_lookup = key already
		sql_cmd = f'SELECT * FROM spec_list WHERE {where_clause}'
		print(sql_cmd, key)
		cur.execute(sql_cmd, key)
		if len(cur.fetchall()) == 0:
			print('The key doesn\'t exist.')
			raise Exception
		for ext in allowed_field[field]:
			if path.endswith(ext):
				md5_hash = self.process_file(path)
				sql_cmd = f'UPDATE spec_list SET {field}=? WHERE {where_clause}'
				print(sql_cmd, (md5_hash,))
				cur.execute(sql_cmd, (md5_hash,) + tuple(key))
				break
		con.commit()
		con.close()

	def spec_doc_to_txt(self, spec_key):
		''' Convert a specification .doc/.docx file to .txt file and 
		attach to txt_file field '''
		con = sqlite3.connect(self.db)
		cur = con.cursor()
		cur.execute(f'SELECT name FROM pragma_table_info("spec_list") \
					WHERE pk <> 0')
		pri_keys = [i[0] for i in cur.fetchall()]
		where_clause = ' AND '.join([f'{pri_key}=?' for pri_key in pri_keys])

		sql_cmd = f'SELECT doc_file FROM spec_list WHERE {where_clause}'
		print(sql_cmd, spec_key)
		doc_file = cur.execute(sql_cmd, spec_key).fetchall()
		if (len(doc_file) == 0 or doc_file[0][0] is None):
			print(f'Doc file doesn\'t exist.')
		elif not os.path.isfile(f'contents/{doc_file[0][0]}'):
			print(f'Doc file was removed, doesn\'t exist anymore.')
		else:
			doc_file = doc_file[0][0]
			self.convert_doc_to_txt(f'contents/{doc_file}', 'tmp/~tmp.txt')
			self.attach_to_spec('tmp/~tmp.txt', spec_key, 'txt_file')
			os.remove('tmp/~tmp.txt')

	def batch_txt_spec(self, *spec_keys):
		''' Convert multiple .doc/.docx files to .txt files '''
		for spec in spec_keys:
			self.spec_doc_to_txt(spec)

	def open(self, key, field):
		''' Open spec file of the key '''
		con = sqlite3.connect(self.db)
		cur = con.cursor()
		cur.execute(f'SELECT name FROM pragma_table_info("spec_list") \
					WHERE pk <> 0')
		pri_keys = [i[0] for i in cur.fetchall()]
		where_clause = ' AND '.join([f'{pri_key}=?' for pri_key in pri_keys])
		sql_cmd = f'SELECT {field} FROM spec_list WHERE {where_clause}'
		print(sql_cmd, key)
		file = con.execute(sql_cmd, key).fetchall()
		if len(file) == 0:
			print(f'Key {key} doesn\'t exist.')
		elif file[0][0] == None:
			print(f'Key {key} exists but the file in {field} not found.')
		else:
			file = f'contents/{file[0][0]}'
			os.startfile(os.path.abspath(file))
	
	def clean_database(self):
		''' Remove all unused files. '''
		pass

	def diff_gen(self, content1, content2, content_type='file', diffmode='eff', \
		output_file='tmp/diff.html', open_output=False):
		''' Function to compare content1 vs content2.
		+ If content_type == 'text', the function directly compares content1 
		vs content2. If content_type == 'file', the function compares 
		contents in path of content1 and content2.
		+ diffmode is the compare mode of diff-match-patch. If diffmode='raw', 
		the function returns diff_main(content1, content2). 
		diffmode='sem' -> diff_cleanupSemantic(diffs)
		diffmode='eff' -> diff_cleanupEfficiency(diffs) with diff_match_patch.Diff_EditCost = 4.
		+ output_file is the destination file of output, formatted as a HTML document.
		+ open_output: Open output file after differenting or not.'''
		
		# Load content1 & content2
		if content_type == 'text':
			text1 = content1
			text2 = content2
		elif content_type == 'file':
			# Read the contents of the text files
			with open(content1, 'r', encoding='utf-8') as file1:
				text1 = file1.read()
			
			with open(content2, 'r', encoding='utf-8') as file2:
				text2 = file2.read()
				
		# Generate the diff between the texts
		dmp_instance = dmp.diff_match_patch()
		dmp_instance.Diff_Timeout = 5.0
		diff = dmp_instance.diff_main(text1, text2)
		
		# Perform cleanup operations on the diff
		if diffmode == 'eff':
			dmp_instance.diff_cleanupEfficiency(diff)
		elif diffmode == 'sem':
			dmp_instance.diff_cleanupSemantic(diffs)
		elif diffmode == 'raw':
			pass
		
		# Create formatted output
		output = ""
		for diff_tuple in diff:
			operation, text = diff_tuple
		
			if operation == -1:
				# Removed text with light red background
				output += '<del style="background-color: #ffe6e6;">' + html.escape(text) + '</del>'
			elif operation == 1:
				# New text with light green background
				output += '<ins style="background-color: #e6ffe6;">' + html.escape(text) + '</ins>'
			else:
				# Unchanged text
				output += html.escape(text)
		
		# Wrap the output in <pre> tags to preserve line breaks
		output = '<pre style="font-family: inherit; white-space: pre-wrap; word-wrap: break-word;">' + output + '</pre>'
		
		# Save output to a file
		with open(output_file, 'w', encoding='utf-8') as diff_file:
			diff_file.write(output)
		
		# Open diff.html in the default web browser
		if open_output == True:
			webbrowser.open(os.path.abspath(output_file))

	def compare_spec(self, spec1, spec2):
		''' Compare 2 specs with key spec1 & spec2 and open differences '''
		con = sqlite3.connect(self.db)
		cur = con.cursor()
		cur.execute(f'SELECT name FROM pragma_table_info("spec_list") \
						WHERE pk <> 0')
		pri_keys = [i[0] for i in cur.fetchall()]
		where_clause = ' AND '.join([f'{pri_key}=?' for pri_key in pri_keys])
		sql_cmd = f'SELECT txt_file FROM spec_list WHERE {where_clause}'
		print(sql_cmd, spec1)
		file1 = cur.execute(sql_cmd, spec1).fetchall()
		print(sql_cmd, spec2)
		file2 = cur.execute(sql_cmd, spec2).fetchall()
		if (len(file1) == 0 or file1[0][0] is None):
			print(f'Text file of specification {spec1} is not avaiable.')
			return None
		elif (len(file2) == 0 or file2[0][0] is None):
			print(f'Text file of specification {spec2} is not avaiable.')
			return None
		else:
			file1 = f'contents/{file1[0][0]}'
			file2 = f'contents/{file2[0][0]}'
			self.diff_gen(file1, file2, open_output=True)