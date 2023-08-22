# cdxToWordTable
Convert SDF to cdx format and insert ChemDraw Object in a Microsoft Word ( docx ) Table

This issue arises yearly for our Chemists as they write patents and every year. They want to import the structure as a chemdraw object in the orientation that it was registered in.  Many tools including the ChemDraw for Excel plugin have issue.  Either they cannot import a structure or if they do, it's essentially a served object.  For example, our chemist converted SMILES to structure so it didn't match registration but when extracted from Excel, it was not actually and OLEObject in the cell when saved.  It was linked to the ChemDraw addin which activated upon opneing.  Not finding any solutions on the web, I explored this problem learning the VBA through Win32 tools to investigate the structure of the desired object.

I've provided a example solution on how to accomplish it.

#### Requirements
1. ChemDraw executable and license.
2. Microsoft Word and license.

### How To
1. This implementation assumes that the files are in `C:/workingDir`.  Change this `dirPath` variable to match the folder that you’re running it in.
2. Make sure the sub folders are in place or change code appropriately: `sdf`, `cdx`, `word`
3. Populate the sdf folder with desired named files with single structures with the filename as the ID.
   - I did this by adding a DB call to get the structures from our registration DB from the list in an Word Table
4. Open a terminal in the working directory and excecute: `python .\cdx_to_Word_Table.py`

