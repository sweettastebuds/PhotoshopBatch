'''
Usage:
	Edits photoshop files

Author:	
	Michael Fernandez

Version:
	1.0.1
'''
from comtypes import client
from xlrd import open_workbook
import os
from time import sleep as sleep

DEBUG = False

# Change dims_path to the root path of Dimentions PSDs
dims_path = r'J:\RM\Bassett\Internal_Dimension_RM\Dimensions\Working\20181112'

# Path to the Dimentions_Setup Doc
dim_setup_doc = r"C:\Users\mfernandez\Desktop\Dimensions Master Sheet.xlsx"

# Change sheet name to the appropriate name
sheet_name = 'Bassett'

# Only updates assets if the status requirement is met.
NeedsDimensions_Only = False
status_requirement = 'NeedsDimensions'

StartAt = 'Lori_Uph_Ottoman_Bench_Tufted'

# Dimensions in Asset Folders
Asset_Folders = True

# Quit Photoshop when the Dimensions Processing has completed.
ExitOnDone = False

# Measurement acronym/denotation
m_type = '\"'

# Change the value of each column if needed
# 0 = A , 1 = B, and so forth
assetname_col = 0
width_col = 1
depth_col = 3
height_col = 4
status_col = 6

# Gets dimensions of an asset
def getDimensions(row):
    # Removes any unnecessary decimal points
    
    w = sheet.cell_value(row, width_col)
    if type(w) is float and w.is_integer():
        w = int(w)
    d = sheet.cell_value(row, depth_col)
    if type(d) is float and d.is_integer():
        d = int(d)
    h = sheet.cell_value(row, height_col)
    if type(h) is float and h.is_integer():
        h = int(h)
    
    return (w,d,h)
    w,h,d = [0,0,0]

# Function containing Photoshop Actions
def PhotoshopCommands(psdFile):
    psApp.Open(psdFile)
    print "Running Commands"
    doc = psApp.Application.ActiveDocument
    layerSets = doc.LayerSets
    
    # Loops through layerSets(aka Photoshop Groups)
    if layerSets.Count > 0:
        for layerSet in layerSets:
                
#            print layerSet.Name
            # Loops through Layers
            for layer in layerSet.ArtLayers:
                sleep(.1)
                # Checks for Text Layers Only
                if layer.Kind == 2:
                    if layerSet.Name.startswith("W") and dimensions[0] != 0:
                        layer.TextItem.contents = str(dimensions[0])+m_type
                        
                    if layerSet.Name.startswith("D") and dimensions[1] != 0:
                        layer.TextItem.contents = str(dimensions[1])+m_type
                        
                    if layerSet.Name.startswith("H") and dimensions[2] != 0:
                        layer.TextItem.contents = str(dimensions[2])+m_type
                        
    doc.Save()                    
    doc.Close(2)                      
    
if __name__ == '__main__':
#    Creates COM object for Photoshop
    psApp = client.CreateObject("Photoshop.Application", dynamic = True)
    psApp.Visible = False
#    psApp = client.GetActiveObject("Photoshop.Application")
    print psApp
    
    # Gets Excel Data
    book = open_workbook(dim_setup_doc, on_demand=True)
    sheet = book.sheet_by_name(sheet_name)

    missing_psd = []
    
    for row in range(sheet.nrows):
        # Skips the Header Row
        if row == 0:
            continue
        
        status = str(sheet.cell_value(row, status_col))
        asset = str(sheet.cell_value(row, assetname_col))
        if asset == u'':
            continue
#        print(asset, status)    #For testing purpose

        elif StartAt:
            start = False
            if not asset in StartAt and start == False:
                continue

            elif asset in StartAt and start == False:
                start = True
                print("Starting from: %s" %asset)
                StartAt = ''
        
        # If NeedsDimentions_Only is True, skips all other statuses
        if NeedsDimensions_Only and status != status_requirement:
            print('Skipping: %s' %asset)
            continue
    
        dimensions = getDimensions(row)
        if dimensions == (u'', u'', u''): continue
        
        #Get PSDs for current Asset
        psdFiles = []
        if not Asset_Folders:
            for psd in os.listdir(dims_path):
                dim_path = dims_path
                #Adds only psd files and filenames matching exact asset name
                if psd.split('-')[0].startswith(asset) and psd.endswith('.psd'):
                    psdFiles.append(psd)
                    
        elif Asset_Folders:
            for assetfolder in os.listdir(dims_path):
                #Searches for an asset folder that matches the asset name
                if asset == assetfolder:
                    dim_path = os.path.join(dims_path, asset)
                    
                    for psd in os.listdir(dim_path):
                        #Adds only psd files and filenames matching exact asset name
                        if psd.endswith('.psd'):
                            psdFiles.append(psd)
                            
        else:
            print('Skipping: %s' %asset)
            continue
        
        if DEBUG:
            print("""DEBUG:
                Dimensions: W %i, D %i, H %i

                """%(dimensions[0],dimensions[1],dimensions[2]))
            continue

        #Loop through PSD files
        print('Processing: %s' %asset)
        if psdFiles == []:
            print('PSD files don\'t exist for: %s' %asset)
            missing_psd.append(asset)

        for psdFile in psdFiles:
            # Contains PSD filepath
            psdFile_path = os.path.join(dim_path, psdFile)
            # Checks if psdFile_path is a file
            if not os.path.isfile(psdFile_path): continue
            PhotoshopCommands(psdFile_path)
    
    # Exits Photoshop
    if ExitOnDone:
        psApp.Quit()
    # Returns a message when completed
    print("Process Complete!")
    print("Missing PSD List:", missing_psd)