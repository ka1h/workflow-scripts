import glob, os, re

rootpath = glob.glob('G:/**/*', recursive=True)

for filename in rootpath:
	new_name = re.sub(r'\(\d{4}_\d{2}_\d{2}\s\d{2}_\d{2}_\d{2}\s.*\)', '', filename)
	os.rename(filename, new_name)
