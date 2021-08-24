# BOM-Export
This addin allows you to export an open assembly to a CSV file so you can manage or import your parts information into other tools or services.

BOM export is a work in progress and as such is not complete. It does, ghowever, fill an important need in its curent state. 

BOM Export supports:

* Flat list of all parts, all levels in the open assembly  
* Removal of sub assemblies so as to create a list of all parts only  
* Part properties as columns  
	* Counts of components: parts and sub assemblies  
	* Component display name as show in browser, including option to supress version number for refrenced designs  
	* Component Description  
	* Components Material. If a components has multiple materials on different bodies, multiple materials will be listed  

##Instalation

###Mac OS
1. Download the main GIT repo [Here.](https://github.com/schneik80/BOM-Export/archive/master.zip)  
2. Extract the zip archive to a convient location.
3. Open a finder window and in the menu bar select **Go to Folder...** Enter "**~/Library**"
![Go to library](./readme-resources/000.png)
4. Browse to **~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/Addins**.
![Addins folder](./readme-resources/001.png)
5. Copy the **BOM-Export** folder you extracted in step 2 into the **Addins Folder**.
![Addins BOM Folder](./readme-resources/002.png)  
6. Next, Launch Fusion 360
7. In the Addins-Pane on the Toolbar select **Scripts and Add-ins...** or press **Shift + S**  
![Addins](./readme-resources/003.png)  
8. Switch to the **Add-Ins** Tab. You should see BOM-Export in the list.
![Addins Tab](./readme-resources/004.png)  
9. Select the **BOM-Export** item in the list and press **Run** and check **Run on startup**. This will ensure that the add-in is always available. You should see BOM-Export now has a running spinner.
![Addins](./readme-resources/005.png)    
10. To confirm that the add-in is available. Close the Add in dialog and look at the bottom of the **Create-Pane**
![Addins](./readme-resources/006.png)  

