# -*- coding: utf-8 -*-
"""
Created on Wed Jan 10 11:05:47 2024

@author: TOBHI
"""
import pandas as pd
import os
import xml.etree.ElementTree as et
from datetime import datetime as dt
import re
import numpy as np
import hashlib


def discover_folders(folderpath, patterns):
    """
    Searches through a folder for subfolders with names that follow the defined
    pattern(s). Used to find project, tower, and flange folders.

    Parameters
    ----------
    folderpath : str
        String of location to search for subfolders in.
    patterns : list
        list of regex pattern expressions to test folder names against.

    Returns
    -------
    folders : TYPE
        list of folders with names matching the given pattern(s).

    """
    # Validate folderpath
    if not folderpath or not os.path.exists(folderpath):
        return []

    # List of all subfolders in directory
    folders = [folder for folder in os.listdir(folderpath)
               if os.path.isdir(os.path.join(folderpath, folder))]
    # Names that will not be used as proj, tower, or flange names
    invalid_names = [".spyproject", "tower_bolt_package",
                     "__pycache__", "Media Files"]

    if len(patterns):
        for folder in folders:
            # Test all patterns for each subfolder
            match = [bool(re.match(pattern, folder.lower()))
                     for pattern in patterns]
            # If none of the patterns match
            if not any(match):
                invalid_names.append(folder)
            else:
                pass
    else:
        pass

    # Return all matching folders
    folders = [folder for folder in folders if folder not in invalid_names]
    return folders


def discover_xmls(folderpath, keyphrase):
    """
    Searches for Xml files in the given folder that have the defined keyphrase
    in either the filename or in the text. Used to determine the file locations
    of the required Xml files.

    Parameters
    ----------
    folderpath : str
        Location of the folder to search for the Xml file in.
    keyphrase : str
        Keyphrase to match the required file. Usually "first" or "second".

    Returns
    -------
    matches : list
        List of file paths for Xmls that match the given keyphrase in the folder.

    """
    # List of all Xmls in directory
    dir_list = [name.lower()
                for name in os.listdir(folderpath) if '.xml' in name]
    matches = []

    for name in dir_list:
        # Iterates through directory list to find required files
        filepath = os.path.join(folderpath, name)

        # Check for file name
        if keyphrase.lower() in name:
            matches.append(filepath)

        # If keyphrase not in filepath, check inside the Xml file
        else:
            with open(filepath) as f:
                f = f.read().lower()
            if "programid" in f and f"installation {keyphrase} round" in f:
                new_filepath = os.path.join(folderpath, f"{keyphrase}_{name}")
                os.rename(filepath, new_filepath)
                matches.append(new_filepath)

    # Return all matching filepaths
    return matches


def parse_round(filepath):
    """
    Parses an Xml file from the smart tensioner tool into dataframes. 
    Expects a known format.

    Parameters
    ----------
    filepath : str
        File location of the Xml file to parse..

    Raises
    ------
    Exception
        DESCRIPTION.

    Returns
    -------
    round_data: Pandas Series
        Pandas Series including dataframes for the header data and the records
        data from the Xml file.

    """
    headers = pd.DataFrame()
    records = pd.DataFrame()
    round_data = pd.Series({"headers": None,
                            "records": None})

    try:
        # Set up XML parsing
        xroot = et.parse(filepath).getroot()

        # Iterate through nodes in xml file
        for nodes in xroot:

            # If node is the header node, build out the header dataframe
            if nodes.tag == 'headers':
                for node in nodes:
                    name = node.find("name").text
                    val = node.find("value").text
                    headers[name] = [val]

            # If this is a records node, fill/build the records dataframe
            elif nodes.tag == 'records':

                # Parse record data into a temporary dict first
                temp_dict = {}
                N_cyc = 0
                for node in nodes:
                    # Iterate through record nodes in bolt
                    name = node.find("name").text   # Name of record
                    if name not in records.columns:  # Add column to dataframe
                        records[name] = ""
                    val = node.find("value").text   # Value of record
                    temp_dict[name] = val           # Push into temporary dict

                    # Get number of bolt cycles
                    if "BoltRotationAngleCycle" in name:
                        n = int(name[22:])  # Number cycle for record node name
                        if n > N_cyc:       # Overwrite if greater
                            N_cyc = n

                # Add cycles to data
                temp_dict["Cycles"] = N_cyc
                if "Cycles" not in records.columns:
                    records["Cycles"] = []

                # Not implemented right now
                # Placeholders for errors to raise if required keys are not found
                if not "BoltNo" in temp_dict.keys():
                    # self.errors += "\nRecord encountered without associated bolt number. Last valid bolt number was #"+records["BoltNo"][-1]
                    # raise Exception(
                    #     "Record encountered without associated bolt number. Last valid bolt number was #"+records["BoltNo"][-1])
                    continue

                else:
                    if not "BoltRotationAngleCycle1" in temp_dict.keys():
                        # self.errors +="\nRequired Key 'BoltRotationAngleCycle1' not found for bolt #"+str(temp_dict["BoltNo"])
                        # raise Exception(
                        #     "Required Key 'BoltRotationAngleCycle1' not found for bolt #"+str(temp_dict["BoltNo"]))
                        pass
                    if not "BoltRotationAngleCycle2" in temp_dict.keys():
                        # self.errors +="\nRequired Key 'BoltRotationAngleCycle2' not found for bolt #"+str(temp_dict["BoltNo"])
                        # raise Exception(
                        #     "Required Key 'BoltRotationAngleCycle2' not found for bolt #"+str(temp_dict["BoltNo"]))
                        pass
                    if not "BoltRotationAngleCycle3" in temp_dict.keys():
                        # self.errors +="\nRequired Key 'BoltRotationAngleCycle3' not found for bolt #"+str(temp_dict["BoltNo"])
                        # raise Exception(
                        #     "Required Key 'BoltRotationAngleCycle3' not found for bolt #"+str(temp_dict["BoltNo"]))
                        pass
                    if not "BoltRotationAngle" in temp_dict.keys():
                        # self.errors +="\nRequired Key 'BoltRotationAngle' not found for bolt #"+str(temp_dict["BoltNo"])
                        # raise Exception(
                        #     "Required Key 'BoltRotationAngle' not found for bolt #"+str(temp_dict["BoltNo"]))
                        pass

                # Push new row into dataframe
                # If bolt number already has a row in the dataframe
                # Use record with more recent date
                bolt_no = temp_dict['BoltNo']
                if any(records.BoltNo.isin([bolt_no])):
                    # Date of first record
                    date1 = dt.strptime(
                        records.loc[bolt_no].Date, "%m/%d/%Y %H:%M:%S")
                    # Date of second record
                    date2 = dt.strptime(
                        temp_dict["Date"], "%m/%d/%Y %H:%M:%S")

                    if date1 > date2:  # First record is newer
                        pass
                    else:  # Second record is newer
                        records.loc[bolt_no] = temp_dict
                else:
                    records.loc[bolt_no] = temp_dict

            else:
                # print("    Unrecognized Node: " + node.tag)
                pass

        # Clear up data from records
        records = records.replace('-', 0)
        records = records.replace('', 0)
        records = records.fillna(0)
        # Fix data types
        for col in records.columns:
            if col == "BoltNo" or col == "Cycles":
                records[col] = records[col].astype(int)
                continue
            else:
                try:
                    records[col] = records[col].astype(float)
                except:
                    pass

        # Push data into return series
        round_data["headers"] = headers.transpose()
        round_data["records"] = records
        return round_data
    except Exception as error:
        print(error)
        return pd.DataFrame()

def find_duplicate_xmls(root_path, search_subfolders=True, criteria="File Name and Size"):
    """
    Searches for XML files and groups them by duplicates based on file size and/or name.
    Similar to how Windows duplicate detection works.

    Parameters
    ----------
    root_path : str
        Root directory path to search for XML files.
    search_subfolders : bool
        If True, search across all subfolders for duplicates.
        If False, only find duplicates within the same folder.
    criteria : str
        Criteria for identifying duplicates:
        "File Name and Size" - Files must have same name AND size
        "File Name Only" - Files with the same name (any size)
        "File Size Only" - Files with the same size (any name)

    Returns
    -------
    list
        List of duplicate groups. Each group is a list of file paths that are duplicates.
        Only includes groups with 2 or more files.
    """
    # Collect all XML files
    all_xml_files = []
    
    for root, dirs, files in os.walk(root_path):
        for file in files:
            if file.lower().endswith('.xml'):
                filepath = os.path.join(root, file)
                try:
                    file_size = os.path.getsize(filepath)
                    all_xml_files.append({
                        'path': filepath,
                        'name': file,
                        'size': file_size,
                        'dir': root
                    })
                except:
                    continue
    
    # Group files based on selected criteria
    duplicates_dict = {}
    
    if search_subfolders:
        # Search across all folders
        for file_info in all_xml_files:
            # Determine key based on criteria
            if criteria == "File Name Only":
                key = (file_info['name'].lower(),)
            elif criteria == "File Size Only":
                key = (file_info['size'],)
            else:  # "File Name and Size"
                key = (file_info['name'].lower(), file_info['size'])
            
            if key not in duplicates_dict:
                duplicates_dict[key] = []
            duplicates_dict[key].append(file_info)
    else:
        # Search only within same folder
        # Group by directory first
        files_by_dir = {}
        for file_info in all_xml_files:
            dir_path = file_info['dir']
            if dir_path not in files_by_dir:
                files_by_dir[dir_path] = []
            files_by_dir[dir_path].append(file_info)
        
        # Find duplicates within each directory
        for dir_path, files in files_by_dir.items():
            for file_info in files:
                # Determine key based on criteria (including dir_path to keep same-folder only)
                if criteria == "File Name Only":
                    key = (file_info['name'].lower(), dir_path)
                elif criteria == "File Size Only":
                    key = (file_info['size'], dir_path)
                else:  # "File Name and Size"
                    key = (file_info['name'].lower(), file_info['size'], dir_path)
                
                if key not in duplicates_dict:
                    duplicates_dict[key] = []
                duplicates_dict[key].append(file_info)
    
    # Filter to only groups with 2+ files and verify content is identical
    duplicate_groups = []
    seen_files = set()  # Track files we've already added to avoid duplicates
    
    for key, file_list in duplicates_dict.items():
        if len(file_list) > 1:
            # Verify files are actually identical by comparing content
            verified_groups = verify_duplicates_by_content(file_list)
            
            # Add each verified group
            for verified_group in verified_groups:
                # Create a frozenset of the group to check if we've seen this exact group before
                group_signature = frozenset(verified_group)
                
                if group_signature not in seen_files and len(verified_group) > 1:
                    duplicate_groups.append(verified_group)
                    seen_files.add(group_signature)
    
    return duplicate_groups


def verify_duplicates_by_content(file_list):
    """
    Verify that files with same name and size actually have identical content.
    
    Parameters
    ----------
    file_list : list
        List of file info dictionaries to verify.
    
    Returns
    -------
    list
        List of lists - each inner list contains file paths that are truly duplicates.
        Returns multiple groups if files have same name/size but different content.
    """
    # Group by content hash
    content_groups = {}
    
    for file_info in file_list:
        try:
            with open(file_info['path'], 'rb') as f:
                content = f.read()
                content_hash = hashlib.md5(content).hexdigest()
                
                if content_hash not in content_groups:
                    content_groups[content_hash] = []
                content_groups[content_hash].append(file_info['path'])
        except:
            continue
    
    # Return all groups that have 2 or more files
    result = []
    for group in content_groups.values():
        if len(group) > 1:
            result.append(group)
    
    return result


# if __name__ == "__main__":
#     test_path = discover_xmls(r'C:\Users\tobhi\OneDrive - Vestas Wind Systems A S\Documents\Git Projects\AME_Engineering\TowerXML\Montgomery Ranch\D04\M2-M3', 'first')
#     test_data = parse_cycle(test_path)
