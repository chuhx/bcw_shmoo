# Basic interface for python to contro Excel 
#
# Revision : $Id: ComExcel.py 1.2 2012/08/14 17:40:53 lchu Exp lchu $
# History : $Log: ComExcel.py $
# History : Revision 1.2  2012/08/14 17:40:53  lchu
# History : automatic check-in
# History :


from win32com.client import Dispatch
import os


class ExcelComObj:

    def __init__(self, sheetnum = None, filename=None):
        self.xlApp = Dispatch('Excel.Application')
        self.exist=0
        self.row_start = 1
        self.col_start = 1
        self.max_rows = 64
        self.max_cols = 6
        self.cell = {'value':'','x1':0 , 'y1' : 0, 'x2':0, 'y2':0}
        self.done = False
        self.row = self.row_start
        self.col = self.col_start
        self.bord_chk = True
        if filename:
            pathname = os.path.split(filename)
            if pathname[0] in [None,'']: path = os.getcwd() + os.sep + "excel"
            else: path = pathname[0]
            nameext = os.path.splitext(pathname[1])
            if nameext[1] in [None,'']: ext = '.xls'
            else: ext = nameext[1]
            name = path + os.sep + nameext[0] + ext
            try:
            	self.xlBook = self.xlApp.Workbooks.Open(name)
            	self.exist=1
            except :
            	self.exist=0
            	print "File does not exist!"
        else:
            self.xlBook = self.xlApp.ActiveWorkbook
            self.filename = ''
        if sheetnum:
        	try:
        		self.sht = self.xlBook.Sheets(sheetnum)
        		self.exist=1
        		self.max_rows = len(self.sht.UsedRange())
        	except :
        		self.exist=0

        else:
        	self.sht = self.xlBook.ActiveSheet



    def getValidCell(self):
		if(self.done):
			return False
		val = None

		for col in range(self.col , self.col_start + self.max_cols):			# CONTINUE FROM THE PREVIOUS SEARCH STOP POINT
			val = self.getCellValue(self.row,col)
			if(val != None):
				row = self.row
				break
					
		if(val == None):
			for row in range(self.row + 1, self.row_start + self.max_rows):		# FIND THE FIRST CELL WHICH HAS VALUE
				for col in range(self.col_start , self.col_start + self.max_cols):
					val = self.getCellValue(row,col)
					if(val != None):
						break
				if(val != None):
					break	
							
		if(val == None):
			self.done = True
			return False
		

		self.row = row
		self.col = col	
		self.cell['value'] = self.getCellText(row,col)
		self.cell['x1'] = self.row
		self.cell['y1'] = self.col

		merged = True
		val = None

		for col in range(self.col + 1 , self.col_start + self.max_cols):	# FIND THE MAX COLUM OF THE CELL 
			val = self.getCellValue(row,col)
			merged = self.checkMerge(row,col)
			if(val != None or merged == False):
				break
				
		if(val == None and merged == True):
			self.col = col
		else:
			self.col = col - 1
		self.cell['y2'] = self.col 
		
		for row in range(self.row + 1 , self.row_start + self.max_rows):	# FIND THE MAX ROW OF THE CELL
			for col in range(self.cell['y1'],self.cell['y2'] + 1):
				val = self.getCellValue(row,col)
				merged = self.checkMerge(row,col)
				if(val != None or merged == False):
					break
			if(val != None or merged == False):
					break
					
		if(val != None or merged == False):
			self.cell['x2'] = row - 1
		else:
			self.cell['x2'] = row

		self.col = self.cell['y2'] + 1										# SET NEW START POINT
		self.row = self.cell['x1'] 
		return(True)

		
    def save(self, filename=None):
        if filename:
            pathname = os.path.split(filename)
            if pathname[0] in [None,'']: path = os.getcwd() + os.sep + "excel"
            else: path = pathname[0]
            nameext = os.path.splitext(pathname[1])
            if nameext[1] in [None,'']: ext = '.xls'
            else: ext = nameext[1]
            name = path + os.sep + nameext[0] + ext
            self.xlBook.SaveAs(name)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, row, col):
		return self.getCellValue(row, col)
          
    def getCellValue(self, row, col):
        suc = False
        while suc == False:
          try:
            val = self.sht.Cells(row, col).Value
            suc = True
          except: pass
        return val
        
    def getCellText(self, row, col):
        suc = False
        while suc == False:
          try:
            val = self.sht.Cells(row, col).Text
            suc = True
          except: pass
        return val
        
    def setCell(self, row, col, value):
        suc = False
        while suc == False:
          try:
            self.sht.Cells(row, col).Value = value
            suc = True
          except: pass
             
    def getCellColor(self, row, col):
        suc = False
        while suc == False:
          try:
            val = self.sht.Cells(row, col).Interior.Color
            suc = True
          except: pass
        return val
  
    def setCellColor(self, row, col,color):
        suc = False
        while suc == False:
          try:
            self.sht.Cells(row, col).Interior.Color = color
            suc = True
          except: pass
        return 
    
    def checkMerge(self, row, col):
        suc = False
        a = False
        while suc == False:
          try:
            if(self.sht.Cells(row,col).MergeCells): a = True
            suc = True
          except: pass
        return a
    
    # lchu, 2011-09-01 11:03:15
    def unmergeCell(self, row, col):
        self.sht.Cells(row,col).MergeCells = False
        return

    def colMerge(self, row1,col1,row2, col2):
        suc = False
        while suc == False:
          try:
            self.sht.Range(self.sht.Cells(row1, col1), self.sht.Cells(row2, col2)).Merge()
            suc = True
          except: pass
        return 

    def borderChk(self, row1,col1,row2, col2):
        suc = False
        while suc == False:
          try:
            val = self.sht.Range(self.sht.Cells(row1, col1), self.sht.Cells(row2, col2)).Borders.Value
            suc = True
          except: pass
        return val

    def borderOn(self, row1,col1,row2, col2):
        suc = False
        while suc == False:
          try:
            self.sht.Range(self.sht.Cells(row1, col1), self.sht.Cells(row2, col2)).Borders.Value = 1
            suc = True
          except: pass
        return 

    def borderOff(self, row1,col1,row2, col2):
        suc = False
        while suc == False:
          try:
            self.sht.Range(self.sht.Cells(row1, col1), self.sht.Cells(row2, col2)).Borders.Value = 0
            suc = True
          except: pass
        return 

    # lchu, 2011-09-05 17:36:28 ------------------------------------------------------------
    def t_comment(self, row, col):
        #self.addComment(row,col, 'cmt test')
        print self.sht.Cells(row, col).Comment.Shape.TextFrame.Characters().Font.Name
        print self.sht.Cells(row, col).Comment.Shape.TextFrame.Characters().Font.Bold
        print self.sht.Cells(row, col).Comment.Shape.Width
        print self.sht.Cells(row, col).Comment.Shape.Height

    def setCommentRectangle(self, row, col, width, height):
        self.sht.Cells(row, col).Comment.Shape.Width = width
        self.sht.Cells(row, col).Comment.Shape.Height = height

    def setCommentFontBoldOn(self, row, col):
        self.sht.Cells(row, col).Comment.Shape.TextFrame.Characters().Font.Bold = True

    def setCommentFontBoldOff(self, row, col):
        self.sht.Cells(row, col).Comment.Shape.TextFrame.Characters().Font.Bold = False
    #-----------------------------------------------------------------------------------------

    def getComment(self, row, col):
        suc = False
        while suc == False:
          if (self.sht.Cells(row, col).Comment != None):
          	try:
          	  val = self.sht.Cells(row, col).Comment.Text()
          	  #self.sht.Cells(row, col).Font.Color = 0xff0000
          	  suc = True
          	except: pass
          else:
          	val = None
          	suc = True
        return val
        
    def addComment(self, row, col,text):
        suc = False
        while suc == False:
          if (self.sht.Cells(row, col).Comment != None):
            self.sht.Cells(row, col).ClearComments()
          else:
            try:
              self.sht.Cells(row, col).AddComment(text)
              #self.sht.Cells(row, col).Font.Color = 0xff0000
              suc = True
            except: pass
        return 
        
    def clearComment(self, row, col):
        suc = False
        while suc == False:
          try:
            self.sht.Cells(row, col).ClearComments()
            #self.sht.Cells(row, col).Font.Color = 0xff0000
            suc = True
          except: pass
        return

    def clearComments(self, row1, col1, row2, col2):
        suc = False
        while suc == False:
          try:
            self.sht.Range(self.sht.Cells(row1, col1), self.sht.Cells(row2, col2)).ClearComments()
            #self.sht.Cells(row, col).Font.Color = 0xff0000
            suc = True
          except: pass
        return
           
    def clearContent(self, row, col):
        suc = False
        while suc == False:
          try:
            self.sht.Cells(row, col).ClearContents()
            #self.sht.Cells(row, col).Font.Color = 0xff0000
            suc = True
          except: pass
        return
  
    def clearContents(self, row1, col1, row2, col2):
        suc = False
        while suc == False:
          try:
            self.sht.Range(self.sht.Cells(row1, col1), self.sht.Cells(row2, col2)).ClearContents()
            #self.sht.Cells(row, col).Font.Color = 0xff0000
            suc = True
          except: pass
        return
    
    def getRange(self, row1, col1, row2, col2):
        suc = False
        while suc == False:
          try:
            val = self.sht.Range(self.sht.Cells(row1, col1), self.sht.Cells(row2, col2)).Value
            suc = True
          except: pass
        return val

    def setRange(self, leftCol, topRow, data):
        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        suc = False
        while suc == False:
            try:
                self.sht.Range(
                    self.sht.Cells(topRow, leftCol),
                    self.sht.Cells(bottomRow, rightCol)
                    ).Value = data
                suc = True
            except: pass


