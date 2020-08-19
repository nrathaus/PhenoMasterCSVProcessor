#!/usr/bin/python3

###
# This Python script reads a provided PhenoMaster CSV files and returns a single excel file with all the data
#  correlated by time

import sys
import csv
from datetime import datetime, date, time, timezone
import xlsxwriter

header = None

if len(sys.argv) < 1:
  print("Please provide one or more filename to be proccessed")
  sys.exit(0)

###
# This returns one workbook - combined - of all the data, in a pivot table
def aggregatedWorkbook(data):
  workbook = xlsxwriter.Workbook('combined.xlsx')
  worksheet = workbook.add_worksheet()

  columnPos = 1
  worksheet.write(1, 0, 'Date')

  relevantColumnCount = 0

  for index in range(len(header)):
    column = header[index]
    if column == '':
      continue

    if ('Date' in column or 
        'Time' in column or 
        'Animal No.' in column or 
        'Box' in column):
      continue

    relevantColumnCount +=1

  animals = data.keys()

  for index in range(len(header)):
    column = header[index]
    if column == '':
      continue

    if ('Date' in column or 
        'Time' in column or 
        'Animal No.' in column or 
        'Box' in column):
      continue

    animalPos = 0
    for animal in animals:
      if 'Dates' in animal:
        continue

      worksheet.write(1, (animalPos * relevantColumnCount) + columnPos, column)
      animalPos += 1

    columnPos += 1

  print("Number of relevantColumnCount: {}".format(relevantColumnCount))

  animalPos = 0
  for animal in animals:
    animal = 'Animal No. ' + animal
    print("Merging from: {} to {}".format(relevantColumnCount * animalPos, relevantColumnCount * (1 + animalPos) ))
    worksheet.merge_range(0, 1 + relevantColumnCount * animalPos, 0, relevantColumnCount * (1 + animalPos), animal)
    animalPos += 1

  date_format = workbook.add_format({'num_format': 'd mm yyyy hh:mm'})

  dateIndex = 0
  for date in data['Dates']:
    worksheet.write_datetime(2 + dateIndex, 0, date, date_format)
    dateIndex += 1
    
  animalPos = 0
  for animal in animals:
    if 'Dates' in animal:
        continue

    print("Handling: {}".format(animal))
    columnPos = 0 # First column is date
    for index in range(len(header)):
      column = header[index]
      if column == '':
        continue

      if ('Date' in column or 
          'Time' in column or 
          'Animal No.' in column or 
          'Box' in column):
        continue

      values = data[ animal ][ column ]
      valueIndex = 0
      for value in values:
        # print("column: {} - value: {}".format(column, value))
        worksheet.write(2 + valueIndex, 1 + columnPos + (relevantColumnCount * animalPos), value) # We skip the first column which is our Animal header and column header, we skip the first column which is Date

        valueIndex += 1

      columnPos += 1

    animalPos += 1

  workbook.close()


###
# This function creates a workbook per column name (skipping those that aren't data)
def workbookPerColumn(dataSet, data, header):
  relevantColumns = []

  for index in range(len(header)):
    column = header[index]
    if column == '':
      continue

    if ('Date' in column or 
        'Time' in column or 
        'Animal No.' in column or 
        'Box' in column):
      continue

    relevantColumns.append( column )

  print("Number of relevantColumns: {}".format(len(relevantColumns)))

  animals = data.keys()

  for relevantColumn in relevantColumns:
    workbook = xlsxwriter.Workbook('column - {} - {}.xlsx'.format(dataSet, relevantColumn))
    print("Column: {}".format(relevantColumn))
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Date')

    date_format = workbook.add_format({'num_format': 'd mm yyyy hh:mm'})

    dateIndex = 0
    for date in data['Dates']:
      worksheet.write_datetime(1 + dateIndex, 0, date, date_format)
      dateIndex += 1

    animalPos = 0
    for animal in animals:
      if 'Dates' in animal:
        continue

      print("Animal: {}".format(animal))
      worksheet.write(0, 1 + animalPos, 'Animal No. {}'.format(animal))

      values = data[ animal ][ relevantColumn ]
      valueIndex = 0
      for value in values:
        # print("column: {} - value: {}".format(column, value))
        worksheet.write(1 + valueIndex, 1 + animalPos, value)
        valueIndex += 1

      animalPos += 1

    workbook.close()


###
# Read the provided CSV files and place the data inside data dict
filenames = sys.argv[1:]
for filename in filenames:
  ###
  # This will hold all the data from the CSV files opened
  data = {}
  data[ 'Dates' ] = []

  print("Opening filename: {}".format(filename))

  dataSet = ''
  header = None

  with open(file=filename, newline='') as csvfile:
    linereader = csv.reader(csvfile, delimiter=',', quotechar='|')

    ###
    # We want to collect all the 'Date' values
    animalName = None
    lastTime = None
    rowIndex = 0
    for row in linereader:
      rowIndex += 1
      if len(row) == 0:
        continue

      if 'Date' in row[0]:
        header = row
        continue

      if header is None:
        continue

      if '' == row[ header.index("Date") ]:
        continue

      if animalName is None:
        animalName = row [ header.index('Animal No.') ]
        print("animalName: {}".format(animalName))
      
      if animalName != row [ header.index('Animal No.') ]:
        break

      dateTime = None

      date_string = "{} {}:{:02d}".format(row[ header.index("Date") ] , row[ header.index("Time") ], 1)
      
      if lastTime == date_string:
        date_string = "{} {}:{:02d}".format(row[ header.index("Date") ] , row[ header.index("Time") ], 31)

      if lastTime and lastTime > date_string:
        print("[210] skipping: {}, smaller than: {} at rowIndex: {}".format(date_string, lastTime, rowIndex))
        continue

      lastTime = date_string

      try:
        dateTime = datetime.strptime( date_string, '%d/%m/%Y %H:%M:%S' )
      except:
        print("date: '{}' is not valid".format( date_string ))
        continue

      data[ 'Dates' ].append(dateTime)

    ###
    # Reopen the file
    csvfile.seek(0)

    ###
    # We now want to collect the data
    lastTime = None
    rowIndex = 0
    header = None
    for row in linereader:
      rowIndex += 1

      if 0 == len(row): # if it has values
        continue

      if rowIndex == 1:
        dataSet = row[0]
        continue
      
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
        print("[261] skipping: {} missing Weight field".format(date_string))
        continue

      if animal not in data:
        lastTime = None
        data[ animal ] = {}

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
        print("[290] skipping: {}, smaller than: {}".format(date_string, lastTime))
        continue

      lastTime = date_string

      try:
        dateTime = datetime.strptime( date_string, '%d/%m/%Y %H:%M:%S' )
      except:
        print("date: '{}' is not valid".format( date_string ))
        continue

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

  workbookPerColumn(dataSet, data, header)

# print("{}".format(data))
