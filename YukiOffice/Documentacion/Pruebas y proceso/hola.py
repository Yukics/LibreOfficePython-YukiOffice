import socket  # only needed on win32-OOo3.0.0
import uno
from Danny.OOo.OOoLib import * 
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
from com.sun.star.awt import Size
 
from com.sun.star.lang import XMain

# get the uno component context from the PyUNO runtime
localContext = uno.getComponentContext()

# create the UnoUrlResolver
resolver = localContext.ServiceManager.createInstanceWithContext(
				"com.sun.star.bridge.UnoUrlResolver", localContext )

# connect to the running office
ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
smgr = ctx.ServiceManager

# get the central desktop object
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

# access the current writer document
model = desktop.getCurrentComponent()
# access the active sheet
active_sheet = model.CurrentController.ActiveSheet

# access cell C4
cell1 = active_sheet.getCellRangeByName("L15")

# set formula
cell1.setFormula("=SUMA(G16:G18)")

