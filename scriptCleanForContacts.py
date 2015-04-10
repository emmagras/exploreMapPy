#!/usr/bin/env python

import xlrd     # Library that processes excel files
import json     # Library for processing / writing JSON

from slugify import slugify  # Library to slugify strings
from pprint import pformat   # Pretty print output

# File to be processed
IMPORT_FILE = 'ntpepInfo.xls' # not .xlsx
OUTPUT_FILE = 'processed_data.json'

def process_data():
    """
    The main processing function.
    """
    # Open the file
    workbook = xlrd.open_workbook(IMPORT_FILE)
    # Datemode is required for processing dates in Excel files
    datemode = workbook.datemode

    worksheet = workbook.sheet_by_name('contacts')

    # Extract headers from first row of worksheet
    # headers = make_headesrs(worksheet)
    # print '\nHeaders are: %s' % pformat(headers)
    # print 'headers.values()',headers.values()

    # getting number or rows and setting current row as 0 -e.g first
    num_rows, curr_row = worksheet.nrows - 1, 0
    # retrieving keys values(first row values)
    # create [u'abb', u'agency', u'firstLast',
    #          u'first', u'Last', u'title',
    #          u'phone', u'email', u'productTypes']
    keys = [x.value for x in worksheet.row(0)]

    # print "keys",keys
    # building dict, the whole JSON object
    data = dict((x, []) for x in keys)

    # iterating through all rows and fulfilling our dictionary
    while curr_row < num_rows:
        curr_row += 1
        for idx, val in enumerate(worksheet.row(curr_row)):
            if val.value.title():
                # this would have been final line in unaltered function
                # followed by print and return data
                data[keys[idx]].append(val.value)
                colValuesList = []
                colValuesList = data[keys[idx]]
                colValuesList.append(val.value)

            # print "colValuesList idx:{0} len:{1} {2}" \
            #    .format(idx,len(colValuesList),colValuesList)

        # Why aren't column indices in their proper order?
        # Liz said proly something goofy in my data.
        # Note: I had to put spaces in cells without values.

        # [0] abb, [1] Last, [2] title,
        # [3] product, [4] agency, [5] firstLast,
        # [6] phone (correct), [7] email (correct), [8] first

        # list not returning the 5 empty strings and that's okay
        # because the other 3 lists won't return them either
        # so these 4 can be zipped
        firstLastList = list(data.values()[5])

        # len 274 without set. len137 with set
        # print "\nfirstLastList", len(firstLastList)
        titleList = list(data.values()[2])

        #len274 without set. len127 set, should be 137,
        # there must be dup titles.
        # print "\ntitleList", len(titleList)

        phoneList = list(data.values()[6])
        # print "\nphoneList", len(phoneList)
        # len274 without set. len135 set, should be 137,
        # there must be dup phones.
        emailList = list(data.values()[7])
        # print "\nemailList", len(emailList)
        # len274 without set. len134 set, should be 137,
        # there must be dup emails.

        abbsList = list(data.values()[0])
        # print "\nabbsList",len(abbsList) #len284 without set. len51 set.
        agencyList = list(data.values()[4])
        # print "\nagencyList", len(agencyList) #len284 without set. len51 set.

    # Goal 1 Completed: create dict from abbsList where item becomes key,
    # separate empty dicts become values.
    # I used a dictionary comprehension instead of .dict()
        # abbsDict = {k: {} for k in abbsList[0:51]}
        # Goal 1 completed. 51 empty dicts being created by this line
        # print "\nabbsDict",abbsDict

    # Goal 2 Completed: insert agency header
    # and agency values into the 51 empty dicts
    # Outside while loop - zipping and then
    # setting of lists made inside while loop, creating dicts

    # this zip associates abb col with agency col; creates proper pairing
    abAg = zip(data.values()[0], data.values()[4])
    # print "abAg", len(abAg)  #len 284
    abAgSets = set(abAg)  # removed dups
    # print "abAgSets", len(abAgSets) #len 51

    # Below is loop that creates dict with agency matched to
    # main state abb key. I'll add contacts list into this dictionary block

    stateDict = {}

    # like saying, to create a key val pair from the
    # zipped and set list of abbs and agencies...
    for abAgPair in abAgSets:
        # position 0 of this zipped list is abb, position 1 is agency.
        # I removed "contactsList": contactsList
        stateDict[abAgPair[0]] = {"agency": abAgPair[1]}
    print "\nstateDict", stateDict

    # skips the 5 empty string states
    agfLtpeList = zip(data.values()[4],
                      data.values()[5],
                      data.values()[2],
                      data.values()[6],
                      data.values()[7])

    # print "agfLtpeList", agfLtpeList #len284
    agFltpeSets = set(agfLtpeList)  # removed dups
    # print "agFltpeSets", agFltpeSets
    # len 142, was 137 before I replaced .strip with .title in line 40.

    # Below is the loop that creates firstLastDicts
    # that will be wrapped by contactsList

    firstLastDict = {}

    # like saying, to create a key val pair from the
    # zipped and set list of ag and 4 cols.
    for agFltpeGroup in agfLtpeList:
        firstLastDict[agFltpeGroup[0]] = {"firstLast": agFltpeGroup[1],
                                          "title": agFltpeGroup[2],
                                          "phone": agFltpeGroup[3],
                                          "email": agFltpeGroup[4]}
    print "\nfirstLastDict", firstLastDict

    # sonia recd using .setdefault method

    wAgDict = {}

    # like saying, for the keys AgencyName
    for names in firstLastDict.keys():
        # like saying, for the vals 501-569-2337,
        # Quality Assurance...,Kevin.Palmer@ahtd.ar.gov,Kevin Palmer,

        for values in firstLastDict[names]:

            # the vals of this new dict will be put into a list.
            # append these names and these values as defined in line above.
            wAgDict.setdefault(values, []) \
                   .append(firstLastDict[names][values])
    print"\nwAgDict", wAgDict

    # if .setdefault doesn't work, try defaultdict,
    #                             or collections.defaultdict,
    #                             or dict.items()

    # or try:
    # data = [('a', 1), ('b', 1), ('b', 2)]

    # d1 = {}
    # d2 = {}

    # for key, val in data:
    #     # variant 1)
    #     d1[key] = d1.get(key, []) + [val]
    #     # variant 2)
    #     d2.setdefault(key, []).append(val)

    # you can't put all the states values into contacts list
    # until there is a tie between the contact and the appropriate state.
    # print "\nfirstLastDict.values()", firstLastDict.values() #len137
    # contacts = firstLastDict.values() # maybe don't rename the list
    # print "\ncontacts", contacts #len137

    # on april 7, see if you can use dict comprehension nesting
    # of curly brackets and square brackets as illustrated below

#     abbsDict = {k: {k: v for v in agencyList[0:51]} for k in abbsList[0:51]}
# +        print "\nabbsDict", abbsDict

    # for contact in agencyList:
    #     = firstLastDict.values()

    # dict.update(dict2)
    # stateDict.update(firstLastDict)
    # print "Value : %s" %  stateDict

    data = stateDict
    return data

    # Could this help?
    # https://github.com/apillalamarri/python_exercises/blob/master/lesson03_csv_to_dict.py
    # Note that I 'faked' the order of my dictionary by using
    # the row numbers as my keys.

    # dict = {}
    # for index, line in enumerate(lines):
    #     single_line_dict = {}
    #     #print zip(headers, line)
    #     for header, element in zip(headers, line):
    #         #print "header is {0} and element is {1}".format(header, element)
    #         single_line_dict[header] = element
    #         #print single_line_dict
    #     dict[index] = single_line_dict
    # return dict

    # Or this?
    # https://docs.python.org/2/library/functions.html#next
    # http://www.tutorialspoint.com/python/file_next.htm
    # for index in range(5):
    #     line = fo.next()
    # print "Line No %d - %s" % (index, line)

    # This line and what it returns on next line is food for thought
    # {x: x**2 for x in (2, 4, 6)}
    # {2: 4, 4: 16, 6: 36}

    # failed block below. transform into a dictionary with the
    # function dict(). but this only returned first letter of each abb.

    # abbs_dict = {}
    # abbs_dict = dict(data.values()[0])
    # print "abbs_dict", abbs_dict

    # failed block below.
    # print "abbs_dict",abbs_dict
    # for key,val in abbsList:
    #     item = key
    #     single_col_dict.update(item)
    # print "single_col_dict",single_col_dict

    # this block below may work for creating firstLast,
    #                                        title,
    #                                        phone,
    #                                        email,
    #                                        productTypes
    #                                        dicts, which may entail .zip( )

    # num_rows = worksheet.nrows - 1
    # curr_row = -1
    # while curr_row < num_rows:
    #     curr_row += 1
    #     row = worksheet.row(curr_row)
    #     rowValuesList = []
    #     rowValuesList.append(row)
    #     # print "\nrowValuesList", rowValuesList

        # single_col_dict = {} # create dict
        # # print "\nzip(keys,rowValuesList",zip(keys,rowValuesList)
        # for key, value in zip(keys,rowValuesList): #looping through zip list, returns keys and values as strings
        #     print "key is {0} and value[0] is {1}".format(key, value[0]) # prints key is abb and value is AL

        # abbVals = []
        # for abb in key,value[0]:
        #     value[0] = abb
        # # print "abb",abb
        # abbVals.append(abb)
        # print "abbVals",abbVals
                # single_col_dict.update(abb)
            # print "single_col_dict",single_col_dict


            # One abandoned attempt to narrow dict to include just rowValuesList[0], aka, 'abb' values
            # if key == 'abb':
            #     single_line_dict[value] = key #returns properties (key, value pairs, i.e., "labels" and "details") as dict
            # print "single_line_dict",single_line_dict




    # return data



                # print "idx,val.value", idx,val.value

                # abbsLists = []
                # if keys[idx] =='abb': # if condition
                #     abbsLists.append(val.value)
                # print "abbsLists",abbsLists
                # # uniqueAbbLists = list(set(abbsLists))
                # # print "uniqueAbbLists", uniqueAbbLists
                # # if val not in abbs: # The in operator can be used to check if an item is present in the

                # single_line_dict = {} # create
                # print "zip(keys,valValuesList",zip(keys,valValuesList) #keys not matching with vals # zip method returns as list

                # attempt to delete dups from dict that has dups
                # result = {}
                # for key,value in data():
                #     if value not in result.values():
                #         result[key] = value
                # print "value", value
                # print "result",result
                # return result



    # print "data",data
    # return data

data = process_data()
# output = []
# output = output.append(data)

# Write the data to JSON
with open(OUTPUT_FILE, 'w') as f: #~do I need to rename OUTPUT_FILE to processed_data?
    json.dump(data, f) # replaced output


# def make_headers(worksheet):
#     """Make headers"""
#     headers = {}
#     cell_idx = 0
#     while cell_idx < worksheet.ncols:
#         cell_type = worksheet.cell_type(0, cell_idx) # ~ cell type 0 is cell with no value, cell type 1 is cell with value
#         cell_value = worksheet.cell_value(0, cell_idx)
#         cell_value = slugify(cell_value).replace('-', '_')
#         if cell_type == 1:
#             headers[cell_idx] = cell_value
#         cell_idx += 1

#     return headers


# def make_values(worksheet): # consider adding 2nd arg, num_rows as iter var here and cell_val line below
#     """Make values"""
#     values = {}
#     cell_idx = 0 # ~ 1 refers to the 2nd column;
#     while cell_idx < worksheet.ncols:
#         cell_type = worksheet.cell_type(0, cell_idx) # ~ cell type 0 is cell with no value, cell type 1 is cell with value
#         cell_value = worksheet.cell_value(1, cell_idx) # ~ don't hardcode a number into first argument, make a var such as row_num or key
#         #cell_value = slugify(cell_value).replace('-', '_')
#         if cell_type == 1:
#             values[cell_idx] = cell_value
#         cell_idx += 1

#     return values


# This allows the script to be run from the command line
if __name__ == "__main__":
    process_data()
