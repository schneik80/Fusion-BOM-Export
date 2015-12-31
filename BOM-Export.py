#Author-Autodesk Inc.
#Description-Etract BOM information from active design.

import adsk.core, adsk.fusion, traceback

showversion = False #show versions in xref component names, default is off
showsubs = False #show the subassemblies in list, for flat BOM default is off. Children are still diplayed this only affects the sub itself

# walk thru the assembly
def walkThrough(bom):
    mStr = ''
    if showsubs == False:
        for item in bom:
            if item['sub'] < 1:
                mStr += '"' + item['name'] + '","' + str(item['pn']) + '","' + str(item['material']) + '",' + str(item['instances']) + '\n'
        return mStr
    if showsubs == True:
        for item in bom:
            mStr += '"' + item['name'] + '","' + str(item['pn']) + '","' + str(item['material']) + '",' + str(item['instances']) + '\n'
        return mStr

#main
def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        design = app.activeProduct
        
        # Make usre we have a desing
        if not design:
            ui.messageBox('No active design', 'BOM Export')
            return

        # Get all occurrences in the root component of the active design
        root = design.rootComponent
        occs = root.allOccurrences
        
        # Gather information about each unique component
        bom = []
        for occ in occs:
            comp = occ.component
            refocc = occ.isReferencedComponent
            occtype = occ.childOccurrences.count
            #print (occtype)
            jj = 0
            for bomI in bom:
                if bomI['component'] == comp:
                    # Increment the instance count of the existing row.
                    bomI['instances'] += 1
                    break
                jj += 1

            if jj == len(bom):
                # Modify the name if versions are OFF and an occurance is an xref
                if (showversion == False and refocc == True):
                    longname = comp.name
                    shortname = " v".join(longname.split(" v")[:-1])
                else:
                    shortname = comp.name
                    
                mat = ''
                bodies = comp.bRepBodies
                for bodyK in bodies:
                    if bodyK.isSolid:
                        mat += bodyK.material.name
                
                # Add this component to the BOM
                bom.append({
                    'component': comp,
                    'name': shortname,
                    'pn' : comp.partNumber,
                    'material' : mat,
                    'instances': 1,
                    'sub': occtype,
                })

        # Display the BOM in the console
        print ('\n')
        print ('Display Name, ' + 'Part Number, ' + 'Material, '+ 'Count')
        print (walkThrough(bom))
         
        # Display the BOM Save Dialog 
        fileDialog = ui.createFileDialog()
        fileDialog.isMultiSelectEnabled = False
        fileDialog.title = "Save BOM"
        fileDialog.filter = 'Text files (*.csv)'
        fileDialog.filterIndex = 0
        dialogResult = fileDialog.showSave()
        if dialogResult == adsk.core.DialogResults.DialogOK:
            filename = fileDialog.filename
        else:
            return
        
        #Write the BOM    
        output = open(filename, 'w')
        output.writelines('Display Name, ' + 'Part Number, ' + 'Material, '+ 'Count\n')
        output.writelines(walkThrough(bom))
        output.close()            
        
        #confirm save
        ui.messageBox( 'Document Saved to:\n' + filename, '', 0, 2)
            
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

