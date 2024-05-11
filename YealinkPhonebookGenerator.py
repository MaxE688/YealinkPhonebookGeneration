from xml.dom import minidom
import xlrd
import glob
import sys
import ctypes

class ToRemotePhonebook:

  #Constructor
  def __init__(self):
    self.dom = minidom.getDOMImplementation()
    self.phoneBookPath = "RemotePhonebook.xml"

  #Formats data from xls file to a Unit object
  def getUnit(self, i, sheet):
    name = sheet.cell(i, 1).value
    phone1 = str(sheet.cell(i, 2).value)
    phone2 = str(sheet.cell(i, 3).value)
    phone3 = str(sheet.cell(i, 4).value)

    if phone1.endswith(".0"):
      phone1 = phone1[:-2]
    if phone2.endswith(".0"):
      phone2 = phone2[:-2]
    if phone3.endswith(".0"):
      phone3 = phone3[:-2]

    if sheet.ncols == 5:
      return Unit(name, phone1, phone2, phone3)
    elif sheet.ncols == 6:
      photo = sheet.cell(i, 5).value
      return Unit(name, phone1, phone2, phone3, photo)
    else:
      ctypes.windll.user32.MessageBoxW(0, "Input file has improper format", "Error: Yealink Phonebook Generator", 1)
      sys.exit("System Exiting")

  #Pulls data from excel file, and formats it to XML
  def outputData(self, xlsDoc, xmlDoc):
    #load xls document
    wb = xlrd.open_workbook(xlsDoc)
    sheet = wb.sheet_by_index(0)

    root = xmlDoc.documentElement
    departments = []
    deptNames = []

    #Organizes data from xls file
    for i in range (1, sheet.nrows):
      deptName = sheet.cell(i, 0).value
      if deptName != "":
        if deptName in deptNames:
          for d in departments:
            if d.name == deptName:
              d.units.append(self.getUnit(i, sheet))
              break
        else:
          deptNames.append(deptName)
          dept = Department(deptName)
          dept.units.append(self.getUnit(i, sheet))
          departments.append(dept)

    #Creates xml tags to store organized xls data
    for dept in departments:
      menuItem = xmlDoc.createElement("Menu")
      menuItem.setAttribute("Name", dept.name)

      for u in dept.units:
        unitItem = xmlDoc.createElement("Unit")
        unitItem.setAttribute("Name", u.name)
        unitItem.setAttribute("Phone1", u.phone1)
        unitItem.setAttribute("Phone2", u.phone2)
        unitItem.setAttribute("Phone3", u.phone3)

        if hasattr(u, "default_photo"):
          if u.default_photo.find(":") > 0:
            unitItem.setAttribute("default_photo", u.default_photo)
          else:
            unitItem.setAttribute("default_photo", "Resource:" + u.default_photo)
        menuItem.appendChild(unitItem)
      root.appendChild(menuItem)

    #Save xml file with pretty formatting
    with open(self.phoneBookPath,"w") as f:
        f.write(xmlDoc.toprettyxml(indent = "\t"))
    ctypes.windll.user32.MessageBoxW(0, "Complete", "Yealink Phonebook Generator", 1)

  #Creates new XML file for data from excel file
  def createXML(self):
    rootNodename = "YealinkIPPhoneBook"
    xmlDoc = self.dom.createDocument(None, rootNodename, None)
    root = xmlDoc.documentElement
    title = xmlDoc.createElement("Title")
    title.appendChild(xmlDoc.createTextNode("Yealink"))
    root.appendChild(title)

    return xmlDoc

#Used to create Department object
class Department:
  def __init__(self, name):
    self.name = name
    self.units = []

#stores the records for each unit in a department
class Unit:
  # def __init__ (self, name, p1, p2, p3, photo):
  def __init__ (self, *args):
    self.name = args[0]
    self.phone1 = args[1]
    self.phone2 = args[2]
    self.phone3 = args[3]
    if len(args) > 4:
      self.default_photo = args[4]

#Checks for proper arguments (XLS file)
if len(sys.argv) > 1:
  xls = sys.argv[1]
else:
  ctypes.windll.user32.MessageBoxW(0, "No input file provided", "Error: Yealink Phonebook Generator", 1)
  sys.exit("System Exiting")

#Drives the program
pb = ToRemotePhonebook()
xml = pb.createXML()
pb.outputData(xls, xml)
