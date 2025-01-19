-- FileRenamer.applescript
-- Advanced file renaming utility with error handling and validation
-- Version: 2.0

use scripting additions

-- Global properties
property defaultPrefix : ""
property defaultStartNumber : 1
property defaultPadding : 2
property maxFilesWarningThreshold : 100

-- Main handler
on run
    try
        -- Get user preferences
        set userPrefs to getUserPreferences()
        if userPrefs is false then return -- User cancelled
        
        set {selectedFolder, basePrefix, startNumber, zeroPadding, sortMethod} to userPrefs
        
        -- Get and validate files
        set fileList to getFilesFromFolder(selectedFolder)
        if length of fileList is 0 then
            display dialog "No files found in the selected folder." buttons {"OK"} default button "OK" with icon stop
            return
        end if
        
        -- Warn if large number of files
        if length of fileList >= maxFilesWarningThreshold then
            set proceed to display dialog "Warning: " & (length of fileList) & " files found. Proceed with renaming?" buttons {"Cancel", "Continue"} default button "Continue" with icon caution
            if button returned of proceed is "Cancel" then return
        end if
        
        -- Sort files based on user preference
        set sortedFiles to sortFiles(fileList, sortMethod)
        
        -- Preview changes
        set previewList to generatePreview(sortedFiles, basePrefix, startNumber, zeroPadding)
        set userConfirmed to previewChanges(previewList)
        if not userConfirmed then return
        
        -- Perform renaming
        set results to performRename(sortedFiles, basePrefix, startNumber, zeroPadding)
        
        -- Show completion summary
        showCompletionSummary(results)
        
    on error errMsg number errNum
        handleError(errMsg, errNum)
    end try
end run

-- Get user preferences through a series of dialogs
on getUserPreferences()
    try
        -- Select folder
        set selectedFolder to choose folder with prompt "Select the folder containing files to rename:"
        
        -- Get base prefix
        set prefixDialog to display dialog "Enter the base prefix for filenames:" default answer defaultPrefix buttons {"Cancel", "Next"} default button "Next" with icon note
        if button returned of prefixDialog is "Cancel" then return false
        set basePrefix to text returned of prefixDialog
        
        -- Get starting number
        set startNumDialog to display dialog "Enter the starting number:" default answer defaultStartNumber buttons {"Cancel", "Next"} default button "Next" with icon note
        if button returned of startNumDialog is "Cancel" then return false
        set startNumber to text returned of startNumDialog as integer
        
        -- Get zero padding
        set paddingDialog to display dialog "Enter the number of digits for padding (e.g., 2 for 01, 02...):" default answer defaultPadding buttons {"Cancel", "Next"} default button "Next" with icon note
        if button returned of paddingDialog is "Cancel" then return false
        set zeroPadding to text returned of paddingDialog as integer
        
        -- Get sort method
        set sortOptions to {"Creation Date", "Modification Date", "Name", "Size"}
        set sortChoice to choose from list sortOptions with prompt "Choose how to sort files:" default items {"Creation Date"}
        if sortChoice is false then return false
        set sortMethod to item 1 of sortChoice
        
        return {selectedFolder, basePrefix, startNumber, zeroPadding, sortMethod}
    on error errMsg number errNum
        handleError(errMsg, errNum)
        return false
    end try
end getUserPreferences

-- Get list of files from folder
on getFilesFromFolder(theFolder)
    tell application "Finder"
        return (files of theFolder whose kind is not "Folder") as alias list
    end tell
end getFilesFromFolder

-- Sort files based on method
on sortFiles(fileList, sortMethod)
    tell application "Finder"
        if sortMethod is "Creation Date" then
            return sort fileList by creation date
        else if sortMethod is "Modification Date" then
            return sort fileList by modification date
        else if sortMethod is "Name" then
            return sort fileList by name
        else if sortMethod is "Size" then
            return sort fileList by size
        end if
    end tell
end sortFiles

-- Generate preview of changes
on generatePreview(fileList, basePrefix, startNum, padding)
    set previewList to {}
    set counter to startNum
    
    repeat with theFile in fileList
        tell application "Finder"
            set oldName to name of theFile
            set fileExt to name extension of theFile
            if fileExt is not missing value then
                set newName to basePrefix & my formatNumber(counter, padding) & "." & fileExt
            else
                set newName to basePrefix & my formatNumber(counter, padding)
            end if
        end tell
        copy {oldName, newName} to end of previewList
        set counter to counter + 1
    end repeat
    return previewList
end generatePreview

-- Show preview and get user confirmation
on previewChanges(previewList)
    set previewText to "Preview of changes:" & return
    
    repeat with i from 1 to 5
        if i > length of previewList then exit repeat
        set oldName to item 1 of item i of previewList
        set newName to item 2 of item i of previewList
        set previewText to previewText & oldName & " -> " & newName & return
    end repeat
    
    if length of previewList > 5 then
        set previewText to previewText & return & "... and " & ((length of previewList) - 5) & " more files"
    end if
    
    set userChoice to display dialog previewText buttons {"Cancel", "Proceed"} default button "Proceed" with icon note
    return button returned of userChoice is "Proceed"
end previewChanges

-- Perform the actual renaming
on performRename(fileList, basePrefix, startNum, padding)
    set results to {succeeded:0, failed:0, errorList:{}}
    set counter to startNum
    
    repeat with theFile in fileList
        try
            tell application "Finder"
                set fileExt to name extension of theFile
                if fileExt is not missing value then
                    set newName to basePrefix & my formatNumber(counter, padding) & "." & fileExt
                else
                    set newName to basePrefix & my formatNumber(counter, padding)
                end if
                set name of theFile to newName
            end tell
            set succeeded of results to (succeeded of results) + 1
        on error errMsg number errNum
            set failed of results to (failed of results) + 1
            copy {original:theFile, errorMessage:errMsg} to end of (errorList of results)
        end try
        set counter to counter + 1
    end repeat
    return results
end performRename

-- Show completion summary
on showCompletionSummary(results)
    set summaryText to "Renaming complete!" & return & return
    set summaryText to summaryText & "Successfully renamed: " & (succeeded of results) & " files"
    
    if failed of results > 0 then
        set summaryText to summaryText & return & "Failed to rename: " & (failed of results) & " files"
    end if
    
    display dialog summaryText buttons {"OK"} default button "OK" with icon note
end showCompletionSummary

-- Format number with leading zeros
on formatNumber(num, padding)
    set numStr to num as string
    set paddedNum to text -padding thru -1 of ("0000000000" & numStr)
    return paddedNum
end formatNumber

-- Error handler
on handleError(errMsg, errNum)
    display dialog "An error occurred:" & return & errMsg & return & return & "Error number: " & errNum buttons {"OK"} default button "OK" with icon stop
end handleError 