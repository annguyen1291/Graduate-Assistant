import os
import csv

directory = os.fsencode("C:/Users/hans/Desktop/Graduate Assistant/Yahoo Finance/")
YahooFianance = "C:/Users/hans/Desktop/Graduate Assistant/YahooFianance.csv"

with open(YahooFianance, 'a') as fout:
	writer = csv.writer(fout, delimiter=',')
	for file in os.listdir(directory):
		print(file)
		filename = os.fsdecode(file)	
		print(filename)
		with open("C:/Users/hans/Desktop/Graduate Assistant/Yahoo Finance/" + filename,'r') as fin:
			reader = csv.reader(fin, delimiter=',')
			for line in reader:
				writer.writerow([filename, line])