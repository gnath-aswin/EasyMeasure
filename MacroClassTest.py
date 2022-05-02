# LSMMacroDemo
# Generate the makro file
# In this tutorial you will learn the basics of yout macro.
# The following commands are fundamental to generate your macro.

from LSMMacroClass import LSMMacroClass # importing class from package 

if __name__ == "__main__":
	 
	# name of the generated macrofile using LSMMacroClass
	newMacroFileName = 'test_depth_roughness_MaackDSi.mcr'

	# open the macro class and send it to a handle
	LSMMC = LSMMacroClass()

	# tell the handle the macro file name
	LSMMC.macroFileName = newMacroFileName

	# now generate the basics of the makro file
	LSMMC.generateFile()

	# start by setting the PATH and NAME for auto saving the data
	LSMMC.comSetAutoSaveSetting(r'D:\Data\Aswin\Automation\Maack_Si', 'test_depth_roughness')

	# *** Define your template prior to your measurements! Note that it is
    # important to tell the Acquisition program that it has to analyze measured
    # data automatically, see the instruction for more details.***
	#LSMMC.comSetReportTemplate(r'D:\Data\Aswin\Automation\Templates\multipleLine_depth_arealRoughness_tiltCorrected_range100_minDepth10.tpl')

	LSMMC.comSetAlignment(r'D:\Data\Aswin\Automation\Maack_Si\alignment.isc')

	# *** You need to know your starting point -
    # Define your x, y coordinates for the starting point or make it zero if you have 
    # set alignment template. This may be the first area 
    # of your stitching (normally top left).  
    # The coordinates are given in MICROMETERS! ***
	# xStart = 250 # µm
	# yStart = -250 # µm

	# *** Here the variable for the report number is set to zero, because it 
    # will be increased later. ***
	repNr = 65

	# *** The following for-loop makes it possible to take measurements at the
    # starting point and in 3 millimeter steps until to 6 millimeters from the
    # starting x-point and y axis. Note that the
    # positions are given in MICROMETERES! ***
	# for iy in range(0,8000,1000):
	# 	for ix in range(0,8000,1000):
	# 		# Tell the LSM that it has to move to another position. 
	# 		LSMMC.comMoveXYZStage(xStart+ix,yStart-iy)

	# 		LSMMC.comSnapshot(1)

	# 		# *** set the stitching area - here 2 by 2 (x / y) fields are measured. Keep in
	#         # mind that the size of the field depends on your objective. ***
	# 		LSMMC.comSetMultiPointArea(1,1)

	# 		# now start your measurements
	# 		LSMMC.com3DExtendedAcquisition()

	# 		# save report as PDF-file
	# 		LSMMC.comSaveReport(repNr,'EXCEL')

	# 		# *** After you have saved your report, you should close the report. If you
	#         # don't close it, and you analyze hundred of test fields, you may imagine
	#         # what will happen to your computer... ***
	# 		LSMMC.comCloseReport()
	# 		repNr = repNr + 1

	LSMMC.comSetReportTemplate(r'D:\Data\Aswin\Automation\Templates\multipleLine_depth_tiltCorrected_range100_minDepth10.tpl')
	
	xStart = 250
	yStart = -8250
	for iy in range(0,8000,1000):
		for ix in range(0,8000,1000):
			# Tell the LSM that it has to move to another position. 
			LSMMC.comMoveXYZStage(xStart+ix,yStart-iy)

			LSMMC.comSnapshot(1)

			# *** set the stitching area - here 2 by 2 (x / y) fields are measured. Keep in
	        # mind that the size of the field depends on your objective. ***
			LSMMC.comSetMultiPointArea(1,1)

			# now start your measurements
			LSMMC.com3DExtendedAcquisition()

			# save report as PDF-file
			LSMMC.comSaveReport(repNr,'EXCEL')

			# *** After you have saved your report, you should close the report. If you
	        # don't close it, and you analyze hundred of test fields, you may imagine
	        # what will happen to your computer... ***
			LSMMC.comCloseReport()
			repNr = repNr + 1
	# changes the objective 1 to 2		
	LSMMC.comChangeRevolver(3) 
	
	LSMMC.closeFile()
	