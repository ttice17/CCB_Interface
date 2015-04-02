import CCBAuth
import xml.etree.ElementTree as ET
import csv
import sys
import win32com.client
import os

class Family(object):
    '''Represents one family unit in CCB '''
    def __init__(self, id):
        self.id = id # Family ID number
        self.members = [] # Family members
        self.name_1 = '' # First name of husband or primary contact if unknown genders
        self.name_2 = '' # First name of wife or spouse if unknown genders
        self.last_name_1 = '' # Last name of husband or primary contact if unknown genders, only used if different last names
        self.last_name_2 = '' # Last name of wife or spouse if unknown genders, only used if different last names
        self.last_name = '' # Family last name if both husband and wife use same last name
        self.salutation_1 = '' # Salutation to be included for name 1
        self.salutation_2 = '' # Salutation to be included for name 2
    def AssignMember(self, person):
        '''Assigns passed person object as a family member'''
        self.members.append(person)
    def BuildNames(self):
        '''Populates name variables'''

        # initialize list to hold adult person objects
        household_adults = []
        # print('----')
        # loop through each member and determine if adult based on criteria that adults are either
        # Primary Contact, Spouse, or Spouse of Member
        # print(len(self.members))
        for family_member in self.members:
            position = family_member.family_position
            # print(position)
            # print(family_member.first_name)
            # print(family_member.last_name)

            if position == 'Primary Contact' or \
                position == 'Spouse' or \
                position == 'Spouse of Member':
                    household_adults.append(family_member)

        '''	choose name_1 and last_name, order of preference is as follows
            Male
            Primary Contact
        '''
        # Populate lists containing relevant family member info
        genders = [adult.gender for adult in household_adults]
        positions = [adult.family_position for adult in household_adults]
        last_names = [adult.last_name for adult in household_adults]
        first_names = [adult.first_name for adult in household_adults]

        # If one of the adults is male assign this person to name 1
        if 'M' in genders:
            # Find location of male adult
            primary_index = genders.index('M')
            for i,adult in enumerate(household_adults):
                if i == primary_index:
                    # Sets male adult name data under name 1
                    # strips extra whitespace from data
                    self.name_1 = adult.first_name.strip()
                    self.salutation_1 = adult.salutation.strip()
                    self.last_name_1 = adult.last_name.strip()
                else:
                    # Set non-male adult name data under name 2
                    # strips extra whitespace from data
                    self.name_2 = adult.first_name.strip()
                    self.last_name_2 = adult.last_name.strip()
                    self.salutation_2 = adult.salutation.strip()
        # If one of the adults is female assign the person to name 2
        elif 'F' in genders:
            # Find location of female adult
            primary_index = genders.index('F')
            for i,adult in enumerate(household_adults):
                if i == primary_index:
                    # Sets female adult name data under name 2
                    # strips extra whitespace from data
                    self.name_2 = adult.first_name.strip()
                    self.salutation_2 = adult.salutation.strip()
                    self.last_name_2 = adult.last_name.strip()
                else:
                    # Set non-female adult name data under name 1
                    # strips extra whitespace from data
                    self.name_1 = adult.first_name.strip()
                    self.last_name_1 = adult.last_name.strip()
                    self.salutation_1 = adult.salutation.strip()
        # If no genders are known assigns name 1 to the Primary Contact
        else:
            # Find location of Primary Contact
            primary_index = positions.index('Primary Contact')
            for i,adult in enumerate(household_adults):
                if i == primary_index:
                    # Sets Primary Contact name data under name 1
                    # strips extra whitespace from data
                    self.name_1 = adult.first_name.strip()
                    self.salutation_1 = adult.salutation.strip()
                    self.last_name_1 = adult.last_name.strip()
                else:
                    # Sets other adult name data under name 2
                    # strips extra whitespace from data
                    self.name_2 = adult.first_name.strip()
                    self.last_name_2 = adult.last_name.strip()
                    self.salutation_2 = adult.salutation.strip()

        # If there are two adults in the household and both have the same last name
        # sets same_last_name_flag
        if (self.last_name_1 == self.last_name_2) or not (self.last_name_1 and self.last_name_2):
            same_last_name = True
            # sets the family last_name to either last_name_1 if it exists (it should), otherwise last_name_2
            if self.last_name_1:
                self.last_name = self.last_name_1
            elif self.last_name_2:
                self.last_name = self.last_name_2
        # if there is only one adult in the household or if two adults with different last names
        # disables same_last_name flag
        # sets family last name to last_name_1
        else:
            same_last_name = False
            self.last_name = ''

        # if both members use the same last name they will appear on one line
        # sets value to conjunction variable so word 'and' will appear in formatted line
        # otherwise value is set to empty and 'and' will be omitted (1 person household, 2 name household)
        if self.name_1 and self.name_2 and same_last_name:
            conjunction = 'and'
        else: conjunction = ''

        # if both adults use same last name only full_name var is set to formatted single line
        # last_name_1 and last_name_2 values set empty
        # uses FormatNames function
        if same_last_name:
            self.full_name = self.FormatNames([self.salutation_1,self.name_1, conjunction, self.salutation_2,self.name_2,self.last_name])
            self.full_name_1 = ''
            self.full_name_2 = ''
        # if different last names then sets individual last name variables using FormatNames function
        else:
            self.full_name = ''
            self.full_name_1 = self.FormatNames([self.salutation_1,self.name_1, self.last_name_1])
            self.full_name_2 = self.FormatNames([self.salutation_2,self.name_2,self.last_name_2])

    def FormatNames(self, names):
        '''Function will take a list of words and concatenate to a line'''
        #initialize empty string
        full_name = ''
        # iterates through each part in the parameter list, if the value is not empty it will concatenate
        # with the full_name string so far
        for name_part in names:
            if name_part:
                full_name = '{} {}'.format(full_name,name_part)

        # removes extra whitespace from name string and returns formatted line
        full_name = full_name.strip()
        return full_name

    def BuildAddress(self):
        '''Uses family object to build address line 1 and line 2'''
        # build list of family member positions
        positions = [member.family_position for member in self.members]
        # find list position of Primary Contact in the members attribute
        primary_index = positions.index('Primary Contact')
        # if the Primary Contact has a full address assign this address as the family address data
        if self.members[primary_index].street and self.members[primary_index].city \
            and self.members[primary_index].state and self.members[primary_index].zip:
                self.street = self.members[primary_index].street.strip()
                self.city = self.members[primary_index].city.strip()
                self.state = self.members[primary_index].state.strip()
                self.zip = self.members[primary_index].zip.strip()
                # concatenate address data
                self.line_1 = self.street
                self.line_2 = "{0}, {1} {2}".format(self.city, self.state, self.zip)
        # if address data is missing will print the last name (last names if 2 name household)
        # of the household missing a valid mailing address
        # if no address data set default values to an empty string
        else:
            if self.last_name:
                print ("Blank Address for {0}".format(self.last_name))
            else:
                print ("Blank Address for {0}/{1}".format(self.last_name_1,self.last_name_2))
            self.line_1 = ""
            self.line_2 = ""

class Person(object):
    '''Generate a person object based on an "individual" tree element'''
    def __init__(self, element,families):
        '''Parse XML to obtain individuals basic data'''
        self.family_id = element.find('family').attrib['id']
        self.id = element.attrib['id']
        salut = element.find('salutation').text
        discard_salut = ['Mr.', 'Mr', 'Mrs.', 'Mrs'] # list of salutations to discard
        # sets salutation variable only if salutation is not discarded according to list
        if salut and salut not in discard_salut:
            self.salutation = salut
        # discards any salutations listed in discard_salut list
        # salutation will be reset to an empty string in these cases
        else:
            self.salutation = ''
        self.first_name = element.find('first_name').text
        self.last_name = element.find('last_name').text
        self.family_position = element.find('family_position').text
        self.gender = element.find('gender').text
        # Address data is collected using GetAddress function
        (self.street, self.city, self.state, self.zip, self.line_1,self.line_2) = self.GetAddress('mailing',element)
        self.CheckFamily(families)
    def GetAddress(self, type, individual):
        '''Parse XML to obtain individuals address data'''
        # defines desired address field tags
        address_fields = ['street_address','city','state','zip','line_1','line_2']
        # initialize address list
        result = []
        # finds each address component in list
        for field in address_fields:
            path = "./addresses/address[@type='{0}']/{1}".format(type,field)
            result.append(individual.find(path).text)
        # returns completed address list
        return(result)
    def CheckFamily(self, families):
        # checks individuals family ID.  If not family object with this ID exists creates a new one
        # and assigns individual
        # uses AssignMember method of family object
        if self.family_id in families:
            families[self.family_id].AssignMember(self)
        # if the individuals family object already exists, just assigns the individual
        else:
            families[self.family_id] = Family(self.family_id)
            families[self.family_id].AssignMember(self)


def SearchListURL(username,password,saved_search):
    '''Calls Searchlist API and selects correct search ID, then builds search URL'''
    # This is the CCB API call to return valid saved searches
    url = 'https://thequestchurch.ccbchurch.com/api.php?srv=search_list'
    # CCBAuth module contains Auth function to handle Basic Authentication
    # pagehandle contains returned data from CCB
    pagehandle = CCBAuth.Auth(username, password, url)
    # read the returns XML and obtain root
    content = pagehandle.read()
    root = ET.fromstring(content)
    # just to searches data in XML
    searches = root.find('./response/searches')
    # iterate saved searches and find the 'search id' if search name matches
    for search in searches:
        if search.find('name').text == saved_search:
            search_id = search.attrib['id']
    # build CCB API string to execute saved search and return
    url = 'https://thequestchurch.ccbchurch.com/api.php?srv=execute_search&id={}'.format(search_id)
    return url

def SearchRequest(username,password,url):
    '''Send search request and return response as the root of an element tree'''
    pagehandle = CCBAuth.Auth(username, password, url)
    content = pagehandle.read()
    return ET.fromstring(content)

def ReadData(individuals,people,families):
    '''iterates each individual in the returned XML and creates a new person object'''
    for individual in individuals:
        person = (Person(individual,families))
        people[person.id] = person


def main():
    with open('config.ini','r') as login_info:
        login_data = login_info.read().splitlines()
        username = login_data[1]
        password = login_data[4]

    saved_search = 'Mailing List'

    # Build url for saved seach
    url = SearchListURL(username,password,saved_search)
    # find root and element tree object of saved search
    root = SearchRequest(username,password,url)
    tree = ET.ElementTree(root)
    # write returned XML to a file this is disabled for unless required for troubleshooting
    # tree.write("search.xml")

    # defines the XML tags of interest
    attributes = ['family','family_position','first_name','last_name']
    address_data = ['street_address','city','state','zip','line_1','line_2']
    # just to individuals tags
    individuals = root.find('./response/individuals')

    # initialize an empty list to contain people missing genders for troubleshooting
    missing_gender = []

    # initialize empty dictionaries for people and family objects
    people = {}
    families = {}
    # call ReadData function to read XML and generate people and family objects
    ReadData(individuals,people,families)

    # define filepath for csv
    directory = os.path.dirname(os.path.abspath(__file__))
    fileName = "MailingList.csv"
    filePath = os.path.join(directory,fileName)

    # open a new csv file to write data
    with open(filePath, 'w', newline='') as csvfile:
        first_line=True
        # iterates through each family in dictionary
        for key, family in families.items():
            # Calls family methods BuildNames and BuildAddress to build name and address strings
            family.BuildNames()
            family.BuildAddress()
            # assign data to write to a list
            # if family shares a last name write appropriate data
            if family.full_name:
                write_data = [family.full_name,'',family.line_1,family.line_2]
            # if 2 name household adds extra line with second name
            else:
                write_data = [family.full_name_1,family.full_name_2,family.line_1,family.line_2]
            # initialize csv writer
            datawriter = csv.writer(csvfile, dialect='excel')
            # writes header row if this is the start of the file
            if first_line:
                datawriter.writerow(['NAME','SECOND NAME','ADDRESS LINE 1','ADDRESS LINE 2'])
                first_line = False
            # write list data to the next row of the csv file
            if write_data[2] and write_data[3]:
                datawriter.writerow(write_data)

    # create new COM object for excel and open CSV file
    # excel = win32com.client.Dispatch("Excel.application")
    # excel.visible = 1
    # excelFile = excel.Workbooks.Open(filePath)
    # # autofit rows to allow adequate width for data
    # excelFile.Worksheets("MailingList").Columns("A:D").AutoFit()

if __name__ == "__main__":
    main()