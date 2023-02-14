# RGC Print Directory Project

## Project initialization

The RGC directory is generated using MS Office template files and macros to import data from an MS Excel-based contact list and images referenced in that list into a MS Word file using its "mail merge" functionality.

The project files are stored in the RGC Staff Google Drive folder [Directory Project Files][1].  To start a new directory follow these instructions.

1. Create a folder at the root of your C: drive called `RGCDirectory`
   > This is VERY important.  All of the templates referrence the folder path `C:\RGCDirectory`.
2. Download all the files in the [Directory Project Files][1] folder and place them into the directory you created in step 1.
   > This will include all the pictures from the previous directory.  You will reuse some of these, replace others and delete any that are no loger needed.

## Data Entry

1. Enter/update directory data into the `RGC_address_list.xlsx` file.
2. Take a photo for each entry, place the photo files into the RGCDirectory folder.
   > We use the naming convention <lastname><firstinitial>.jpg for the image files.
3. Update each entry to have a photo filename in the `Photo` collumn.
4. Edit/normalize the photos using an image editor of your choice .(recommendation: [Paint.NET][2])
   1. Each photo should first be changed to have a resolution of 240 pixels per inch.
   2. Each photo should be edited through a combination of cropping and resizing so that you have an image that is 720px by 720px (3in x 3in).
      1. A common means to this end:
         1. Crop the photo to remove excess backgroound, but include all of all the people.
         2. Resize the photo so that the shorter edge is 720px.
         3. Crop the photo again to its 720px x 720px final desired size.
      2. Sometimes step three cannot be done because the height will be less than 720px in order to get everyone in.  That is OK, but should be avoided when possible.
   3. If the source photo is already LESS than 240 pixels per inch then the important thing here is to make sure that the natural dimensions of the image are 3"x3".  Multiply the resolution by 3 to know what your target image size is. 
      > In the instructions above 240 * 3 is 720.

      > The goal here is to prevent MS Word from needing to manipulate the images in any way.  Either it will not do so, or it will do so badly resulting in images that do not fit or in images that are blury.

## Generate the document

1. Open the `RGC_directory_template.docm` file.
2. Answer yes to all prompts about importing data from the spreadsheet.
3. Update the date in the footer of the document and save it.
4. On the `Mailings` menu select `Finish & Merge` >> `Edit Individual Documents...`.
5. In the `Merge to new document` dialog, sellect 'All' and press [Ok].
6. In the generated document press `[ctrl]-A` to select the entire document.
7. Press `F9` to update all fields, this will pull in all of the images.
8. If prompted with a security warning telling you that the document contains fields that can share data, this is ok, answer `[Yes]` .
9. Update the document as needed.
   1.  Change the names of non-members to be not in bold font.
   2.  Change the names of children who are members to be in bold.
   3.  Add whitespace and a section title for out of town prople at the end.
   4.  Etc.
10. Save the generated document as a PDF document.
11. Print the PDF copy onto regular white paper.

## Print front pages

1. Open the front pages document named `FILENAME`.
2. Update the content as needed.
3. Print onto regular white paper.

## Print the cover

1. Open the cover document named `FILENAME`.
2. Update the publishing date.
3. Update other content as needed.
4. Print onto colored card stock.

## Construct the booklets

1. Take 1 copy of each document.
2. Stack them in the correct order.
3. Staple.

## Project shutdown

Once the directories are completed do the following to ensure that everything is set for the next diretory.

1. Upload the final PDF version to the [Directories][3] folder in Google Drive.
2. Delete all of the files from the [Directory Project Files][1] folder in Google Drive.
3. Upload the entire content of `C:\RGCDirectory` on your computer to the [Directory Project Files][1] folder in Google Drive.  This should include
   1. The contact list Excel spreadsheet (`RGC_address_list.xslx`).
   2. The directly template file (`RGC_directory_template.docm`).
   3. The directory cover file (`FILENAME`).
   4. The fromt pages document (`FILENAME`).
   5. All the image files used for this version.
   6. This instruction document (`README.md`)


[1]: https://drive.google.com/drive/u/0/folders/13o3A3-BaKan7npegzg7bu4kf5i3uJAT2
[2]: https://www.getpaint.net/download.html
[3]: https://drive.google.com/drive/u/0/folders/1oNXtEiJhaJemut3YAPWG4JV7eigZC7zL
