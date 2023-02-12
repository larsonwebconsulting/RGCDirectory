# RGCDirectory

## Installation

1. Create a folder at the root of your C: drive called `RGCDirectory`
2. Download all the files in this repository and place them into the directrory you created in step 1

## Data Entry

1. Enter/update directory data into the RGC_address_list.xlsx file
2. Take a photo for each entry, place the photo files into the RGCDirectory folder
3. Update each entry to have a photo filename in the `Photo` collumn.
4. Edit/normalize the photos using an image editor of your choice (recommendation: [Paint.NET](https://www.getpaint.net/download.html))
   1. Each photo should first be changed to be 240 pixels per inch
   2. Each photo should be edited through a combination of cropping and resizing so that you have an image that is 720px by 720px (3in x 3in).
      1. A common means to this end:
         1. Crop the photo to remove excess backgroound, but include all of all the people.
         2. Resize the photo so that the shorter edge is 720px
         3. Crop the photo again to its 720px x 720px final desired size.
      2. Sometimes step three cannot be done because the height will be less than 720px in order to get everyone in.  That is OK, but should be avoided when possible.

## Generate the document

1. Open the RGC_directory_template.docm file
2. Answer yes to all prompts about importing data from the spreadsheet
3. Update the date in the footer of the document and save it.
4. On the `Mailings` menu select `Finish & Merge` >> `Edit Individual Documents...`
5. In the `Merge to new document` dialog, sellect 'All' and press [Ok]
6. In the generated document press `[ctrl]-A` to select the entire document
7. Press `F9` to update all fields, this will pull in all of the images.
8. If prompted with a security warning telling you that the document contains fields that can share data, this is ok, answer `[Yes]` 
9. Save the generated document as a PDF document