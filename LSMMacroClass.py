class LSMMacroClass:

	def __init__(self):
		self.macroFileName = ''
		self.fid = ''
		self.ASFolderName = ""
		self.ASFileName = ""

	#copy the contents of the .bak file to create array
	#line numbering starts from 0
	#variable content is an array 
	def importDummyMacro(self):
		self.name = 'DummyMakro.bak'
		self.dummyFile = open(self.name,"rt")
		self.content = self.dummyFile.readlines()
	
	#copy required lines with index to macro file
	def copyDummyLines(self, nStart, nEnd):
		self.lines = self.content[nStart:nEnd]
		for line in self.lines:
			self.generate.write(line)

	
	#write the required modification to macro commands other than copying standard code
	def writeLine(self,text):
		self.generate.write(''.join(text))
		self.generate.write('\n')
		pass

	def generateFile(self):
		self.importDummyMacro()
		self.generate = open(self.macroFileName, 'w')
		self.copyDummyLines(0,4)

	def closeFile(self):
		self.copyDummyLines(519,520)
		self.generate.close()

	#commands copied from DummyMakro.bak -> they should be in this order in the DummyMakro File 
	#makes it easier if the list gets extended

	def comSetAutoSaveSetting(self, folderName, fileName):
		self.ASFolderName = folderName
		self.ASFileName = fileName
		self.copyDummyLines(4,10)
		self.writeLine(['        <value>'+fileName+'</value>'])
		self.copyDummyLines(11,17)
		self.writeLine(['        <value>'+folderName+'</value>'])
		self.copyDummyLines(18,72)
		pass

	def comSetReportTemplate(self, fileName):
		self.copyDummyLines(72,78)
		self.writeLine(['        <value>'+fileName+'</value>'])
		self.copyDummyLines(79,91)
		pass

	def comSetMultiPointArea(self, nCol, nRow):
		self.copyDummyLines(91,111)
		self.writeLine(['        <value>'+str(nCol)+'</value>'])
		self.copyDummyLines(112,118)
		self.writeLine(['        <value>'+str(nRow)+'</value>'])
		self.copyDummyLines(119,145)
		pass

	def comMoveXYZStage(self, xPos, yPos):
		self.copyDummyLines(145,151)
		self.writeLine(['        <value>'+f"{xPos:.3f}"+ '</value>'])
		self.copyDummyLines(152,158)
		self.writeLine(['        <value>'+f"{yPos:.3f}"+'</value>'])
		self.copyDummyLines(159,178)
		pass

	def com3DExtendedAcquisition(self):
		self.copyDummyLines(178,184)
		self.writeLine(['        <value>'+self.ASFileName+'</value>'])
		self.copyDummyLines(185,191)
		self.writeLine(['        <value>'+self.ASFolderName+'</value>'])
		self.copyDummyLines(192,330)
		pass

	def comSaveReport(self, N, ext):
		self.copyDummyLines(330,336)
		self.writeLine(['        <value>'+self.ASFolderName+'</value>'])
		self.copyDummyLines(337,343)
		self.writeLine(['        <value>'+self.ASFileName+'_'+f"{N:04}"+'</value>'])
		self.copyDummyLines(344,350)
		self.writeLine(['        <value>'+ext.upper()+'</value>'])
		self.copyDummyLines(351,370)
		pass

	def comCloseReport(self):
		self.copyDummyLines(410,429)
		pass

	def comSnapshot(self, N):
		self.copyDummyLines(429,435)
		self.writeLine('        <value>'+self.ASFileName+'_'+f"{N:04}"+'</value>')
		self.copyDummyLines(436,442)
		self.writeLine('        <value>'+self.ASFolderName+'</value>')
		self.copyDummyLines(443,462)
		pass


	def comChangeRevolver(self, N):
		if N in range(1,4):
			self.copyDummyLines(462,468)
			self.writeLine('        <value>'+str(N)+'</value>')
			self.copyDummyLines(469,488)
		else:
			print('Revolver can only be change to Position 1,2 or 3')
		pass

	def comSetAlignment(self, fileName):
		self.copyDummyLines(488,494)
		self.writeLine('        <value>'+fileName+'</value>')
		self.copyDummyLines(495,519)
		pass











