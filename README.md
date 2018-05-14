# battletechTools
Tool to extract json data to Excel, edit in Excel, and then push back to json.


Main path of json files: BATTLETECH\BattleTech_Data\StreamingAssets\data

# To use this tool:
Create a subdirectory from the location of main.py named "data".

Then, put any directories from your "BATTLETECH\BattleTech_Data\StreamingAssets\data" 
path that you'd like to edit inside the new data directory.

Run main.py

Option 0 will locate a "movement" directory inside the new data folder, 
and update any json keys that are tied to movement velocity. This will not 
change the distance any mech/vehicle moves per turn.

Option 1 will pull any json data from a specified subdirectory inside the new data 
folder, and put it into an excel file named "jsonData.xlsx" in your main path.

Option 2 will push data from jsonData.xlsx to an "export" directory. To use these 
files, move/copy them to their appropriate path in your BattleTech_Data folder.

# Disclaimer
This is a slapped-together tool. I hope it's useful for some folks. If not, 
at least it's been a good learning exercise for learning python and using github.
