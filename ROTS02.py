#import for RegOnline piece
from pysimplesoap.client import SoapClient
from pysimplesoap.simplexml import SimpleXMLElement
import re
import xml.etree.cElementTree as ET

#added out of place to make sure that later datetime imports work
import datetime

#more imports for RegOnline
from datetime import datetime
from datetime import timedelta
import pymssql
import array
import sys
import xmltodict

#import for SharePoint piece

from suds.client import Client
from suds.transport.https import WindowsHttpAuthenticated
from suds.sax.element import Element
from suds.sax.attribute import Attribute
from suds.sax.text import Raw

def GetNow():
	return str(datetime.now())

def initializeReg(apitoken):
	client = SoapClient(
		wsdl = "https://www.regonline.com/api/default.asmx?WSDL"
		, trace = False)

	header = SimpleXMLElement("<Headers/>")

	MakeHeader = header.add_child("TokenHeader")
	MakeHeader.add_attribute('xmlns','http://www.regonline.com/api')
	MakeHeader.marshall('APIToken', apitoken)
	client['TokenHeader']=MakeHeader
	return client
	
def LastSixMonths():
	#set the filter for time; set to anything changed (modified) in the last 180 days
	SixMonthsAgo = datetime.now() - timedelta(days=180)
	SixMonthsAgoString = SixMonthsAgo.strftime("%Y,%m,%d")
	FilterTime = "ModDate >= DateTime(" + SixMonthsAgoString + ")"
	return FilterTime
	
def ProcessRegData(ResponseString):
	#process the ResponseString from RegOnline into an xml object and return it
	#does this by finding the <Data> tags and cutting off prefix and suffix
	FindDataStart = ResponseString.find("<Data>")
	FindDataEnd = ResponseString.find("</Data>")
	DataEnd = FindDataEnd + 7
	#get all the data between the <Data> tags
	ReducedData = ResponseString[FindDataStart:DataEnd]
	#replace the string with nothing ""
	root = ReducedData.replace('xsi:nil="true"',"")
	return root
	
def MakeRootDict(root):
	#parses the xml returned from RegOnline into a Dictionary
	User_Info = xmltodict.parse(root)
	return User_Info
	
def initializeSP(SPurl, SPusername, SPpassword):

	url= SPurl + '_vti_bin/lists.asmx?WSDL'
	ntlm = WindowsHttpAuthenticated(username=SPusername, password=SPpassword)
	client = Client(url, transport=ntlm)
	return client
	
def writeUsers(User_Info, client, StudentInfo):
	#set tallies to zero
	UpdatedRecords = 0
	NewRecords = 0
	for CurrentItem in User_Info['Data']['APIRegistration']: 
		item_data = {}

		for DataPoint in StudentInfo:
			item_data.update({DataPoint:CurrentItem[DataPoint]})
	
		#set the RegOnlineID to the same thing as the "id" that you get from RegOnline, then delete the "ID"
		#can't set the ID column in SharePoint through SOAP services. idk, I guess because it's a system value?
		item_data["RegOnlineID"] = item_data['ID']
		del item_data['ID']
		
		#Set a blank variable for the GetListItems request
		blank = ""
		#setup the xml query for checking for the ID number
		Eq = Element('Eq')
		Eq.append(Element('FieldRef').append(Attribute('Name','RegOnlineID')))
		Eq.append(Element('Value').append(Attribute('Type','Text')).setText(item_data["RegOnlineID"]))
		Where = Element('Where')
		Where.append(Eq)
		Query = Element('Query')
		Query.append(Where)
		query = Query
	                                                
		responseExist = client.service.GetListItems('{60848478-FFC6-4897-81BE-C956C55A9B10}', blank, Raw(query))
		
		#print responseExist
	
		if responseExist.listitems.data._ItemCount == "1":
	
			#set item id to returned id from GetListItems
			#print responseExist.listitems.data.row
			item_data["ID"] = responseExist.listitems.data.row._ows_ID
		
			#Begin creating the updates item by defining a batch
			batch = Element( 'Batch' )
			batch.append(Attribute('OnError','Continue')).append(Attribute('ListVersion','1'))
	
	
			#second level element needed to update. notice the Update attribute for the Cmd
			#left all the options in here just in case I wanted to use them in the future.
			method = Element( 'Method')
			method.append(Attribute('ID','1')).append(Attribute('Cmd','Update'))
			#method.append(Attribute('ID','1')).append(Attribute('Cmd','New'))
			#method.append(Attribute('ID','1')).append(Attribute('Cmd','Delete'))
			#method.append(Attribute('ID','1')).append(Attribute('Cmd','Move'))
	
			#add a field for every dictionary item
			for key in item_data:
				val = item_data[ key ]
				#get rid of spaces in column names
				key = key.replace(' ','_x0020_')
				#correct date to format
				#if isinstance( val, datetime.datetime):
				if hasattr(val,"datetime"):
					val = datetime.datetime.strftime(val, '%Y-%m-%d %H:%M:%S')
				method.append( Element('Field').append(Attribute('Name', key)).setText(val))
	
			#add method object into the batch object
			batch.append(method)
	
			#set the name as updates for the way suds formats the xml
			updates = batch
		
			try:
				response = client.service.UpdateListItems('{60848478-FFC6-4897-81BE-C956C55A9B10}', Raw(updates) )
		
			except Exception as e:
				print str(e)
		
			except suds.webfault as e:
				print str(e)
	
			else:
				#print response.Results.Result.ErrorCode
				#print sys.exc_info()
				print "Record " + item_data["ID"] + " updated"
				UpdatedRecords += 1
	
		elif responseExist.listitems.data._ItemCount > "1":
			#had to add this in case a record made it into the list more than once... might error check it someday...  :/
			continue
		else:
		
			#add this record as a new item to the list
			#Begin creating the updates item by defining a batch
			batch = Element( 'Batch' )
			batch.append(Attribute('OnError','Continue')).append(Attribute('ListVersion','1'))
		
			#second level element needed to update. notice the 'New' attribute for the Cmd
			method = Element( 'Method')
			method.append(Attribute('ID','1')).append(Attribute('Cmd','New'))
		
			#add a field for every dictionary item
			for key in item_data:
				val = item_data[ key ]
				#get rid of spaces in column names
				key = key.replace(' ','_x0020_')
				#correct date to format
				#if isinstance( val, datetime.datetime):
				#	val = datetime.datetime.strftime(val, '%Y-%m-%d %H:%M:%S')
				if (key == 'StartDate') or (key == 'EndDate'):
					if not val is None:
						val = val.replace('T', ' ')
				method.append( Element('Field').append(Attribute('Name', key)).setText(val))
		
			#add method object into the batch object
			batch.append(method)
		
			#set the name as updates for the way suds formats the xml
			updates = batch
			
			try:
				response = client.service.UpdateListItems('{60848478-FFC6-4897-81BE-C956C55A9B10}', Raw(updates) )
		
			except Exception as e:
				print str(e)
		
			except suds.webfault as e:
				print str(e)
		
			else:
				#print response.Results.Result.ErrorCode
				#print sys.exc_info()
				print "Record " + item_data["ID"] + " added to the List"
				NewRecords += 1
	return (UpdatedRecords,NewRecords)

def writeEvents(Event_Info, client, EventInfo):
	#set tallies to zero
	UpdatedRecords = 0
	NewRecords = 0
	for CurrentItem in Event_Info['Data']['APIEvent']: 

		item_data = {}

		for DataPoint in EventInfo:
			item_data.update({DataPoint:CurrentItem[DataPoint]})
	
		#set the RegOnlineID to the same thing as the "id" that you get from RegOnline
		item_data["EventID"] = item_data['ID']
		del item_data['ID']
		
		#Set a blank variable for the GetListItems request
		blank = ""
		#setup the xml query for checking for the ID number
		Eq = Element('Eq')
		Eq.append(Element('FieldRef').append(Attribute('Name','EventID')))
		Eq.append(Element('Value').append(Attribute('Type','Text')).setText(item_data["EventID"]))
		Where = Element('Where')
		Where.append(Eq)
		Query = Element('Query')
		Query.append(Where)
		query = Query
	                                                 
		responseExist = client.service.GetListItems('{12C63117-18E2-4D92-9C0D-38202F86337C}', blank, Raw(query))
		#print responseExist[0][0][1]
	
		if responseExist.listitems.data._ItemCount != "0":
	
			#set item id to returned id from GetListItems
			item_data["ID"] = responseExist.listitems.data.row._ows_ID
		
			#Begin creating the updates item by defining a batch
			batch = Element( 'Batch' )
			batch.append(Attribute('OnError','Continue')).append(Attribute('ListVersion','1'))
	
	
			#second level element needed to update. notice the Update attribute for the Cmd
			method = Element( 'Method')
			method.append(Attribute('ID','1')).append(Attribute('Cmd','Update'))
				
			#add a field for every dictionary item
			for key in item_data:
				val = item_data[ key ]
				
				#get rid of spaces in column names
				key = key.replace(' ','_x0020_')
				#correct date to format
				#if isinstance( val, datetime.datetime):
				if (key == 'StartDate') or (key == 'EndDate'):
					if not val is None:
						val = val.replace('T', ' ')
				method.append( Element('Field').append(Attribute('Name', key)).setText(val))
	
			#add method object into the batch object
			batch.append(method)
			#set the name as updates for the way suds formats the xml
			updates = batch
		
			try:
				response = client.service.UpdateListItems('{12C63117-18E2-4D92-9C0D-38202F86337C}', Raw(updates) )
				#print response
		
			except Exception as e:
				print str(e)
			except suds.webfault as e:
				print str(e)
			else:
				#print response #for troubleshooting
				#print sys.exc_info()
				print "Event " + item_data["ID"] + " updated"
				UpdatedRecords += 1
	
		else:
		
			#add this record as a new item to the list
			#Begin creating the updates item by defining a batch
			batch = Element( 'Batch' )
			batch.append(Attribute('OnError','Continue')).append(Attribute('ListVersion','1'))
		
		
			#second level element needed to update. notice the 'Update' attribute for the Cmd
			method = Element( 'Method')
			method.append(Attribute('ID','1')).append(Attribute('Cmd','New'))
					
			#add a field for every dictionary item
			for key in item_data:
				val = item_data[ key ]
				#get rid of spaces in column names
				key = key.replace(' ','_x0020_')
				#correct date to format
				if (key == 'StartDate') or (key == 'EndDate'):
					if not val is None:
						val = val.replace('T', ' ')
				method.append( Element('Field').append(Attribute('Name', key)).setText(val))
		
			#add method object into the batch object
			batch.append(method)
		
			#set the name as updates for the way suds formats the xml
			updates = batch
			print updates
			
			try:
				response = client.service.UpdateListItems('{12C63117-18E2-4D92-9C0D-38202F86337C}', Raw(updates) )
		
			except Exception as e:
				print str(e)
			except suds.webfault as e:
				print str(e)
			else:
				#print response #for troubleshooting
				#print sys.exc_info()
				print "Event " + item_data["ID"] + " added to the List"
				NewRecords += 1
	return (UpdatedRecords,NewRecords)
