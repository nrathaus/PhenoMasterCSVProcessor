#!/usr/bin/python3

###
# This Python script reads a provided PhenoMaster CSV files and returns a single excel file with all the data
#  correlated by time

import sys
import csv
from datetime import datetime, date, time, timezone
import xlsxwriter

header = None

data = {}

lastTime = None
fakeSeconds = 30

if len(sys.argv) < 1:
  print("Please provide one or more filename to be proccessed")
  sys.exit(0)

filenames = sys.argv[1:]
for filename in filenames:
  print("Opening filename: {}".format(filename))
  header = None

  with open(file=filename, newline='') as csvfile:
    linereader = csv.reader(csvfile, delimiter=',', quotechar='|')
    for row in linereader:
      if len(row) > 0:

        # Date,Time,Animal No.,Box,Ref.SFlow,Ref.O2,Ref.CO2,VO2(3),VCO2(3),RER,H(3),XT+YT,XA,YA,Drink,Feed,Weight,
        if 'Date' in row[0]:
          header = row
          continue

        if header is None:
          continue

        # print("header: {}".format(header))

        # Make sure its an Animal 
        if 'Animal No.' not in header:
          continue

        # Make sure its a proper section of data beacuse PhenoMaster has two sections
        if 'Weight' not in header:
          continue

        animal = row[ header.index('Animal No.') ]
        if '' == animal:
            continue

        # print("Animal No.: {}".format(animal))
        # Sometimes lines are malformed, make sure that the Drink field is there
        indexWeight = header.index('Weight')
        if indexWeight > len(row):
          print("skipping: {} missing Weight field".format(date_string))
          continue

        if animal not in data:
            lastTime = None
            data[ animal ] = {}
            data[ animal ][ 'Date' ] = []

            for column in header:
              if column == '':
                continue

              if ('Date' in column or 
                  'Time' in column or 
                  'Animal No.' in column or 
                  'Box' in column):
                continue
              data[ animal ][ column ] = []

        
        if '' == row[ header.index("Date") ]:
            continue

        dateTime = None

        date_string = "{} {}:{:02d}".format(row[ header.index("Date") ] , row[ header.index("Time") ], 1)
        
        if lastTime == date_string:
            date_string = "{} {}:{:02d}".format(row[ header.index("Date") ] , row[ header.index("Time") ], 31)    

        if lastTime and lastTime > date_string:
            print("skipping: {}, smaller than: {}".format(date_string, lastTime))
            continue

        lastTime = date_string

        try:
            dateTime = datetime.strptime( date_string, '%d/%m/%Y %H:%M:%S' )
        except:
            print("date: '{}' is not valid".format( date_string ))
            continue

        data[ animal ][ 'Date' ].append( dateTime )

        for column in header:
          if column == '':
            continue

          if ('Date' in column or 
              'Time' in column or 
              'Animal No.' in column or 
              'Box' in column):
            continue

          dataValue = row[ header.index( column ) ]
          if dataValue == '-':
            data[ animal ][ column ].append( dataValue )
          else:
            data[ animal ][ column ].append( float( dataValue ) )

        #if len(data[ animal ]['Drink']) > 10:
        #    break
        # print("{}".format( row[ header.index("Drink") ]))

# print("{}".format(data))
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(1, 0, 'Date')

columnPos = 0
for index in range(len(header)):
  column = header[index]
  if column == '':
    continue

  if ('Date' in column or 
      'Time' in column or 
      'Animal No.' in column or 
      'Box' in column):
    continue

  columnPos += 1
  worksheet.write(1, columnPos, column)

animalPos = 0
animals = data.keys()
for animal in animals:
  animal = 'Animal No. ' + animal
  worksheet.merge_range(0, columnPos * animalPos, 0, columnPos * ( 1 + animalPos ), animal)
  animalPos += 1

workbook.close()
