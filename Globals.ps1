#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------

<# 
TODO:
	* Create a new "Save" button.
	* Remove the Save-XML function call from other functions.
	* Save-XML should only be called when you press the 'Save' button.
	
CONSIDERATIONS:
	* Have the changes be saved to a temporary file, instead of only in memory.
<#
	This function will save the XML file.
#>
function Save-XML()
{
	Write-Host "Saving XMl."
	$ManagementPackXml.Save($ManagementPackPath)
}

<#
	This function will query the XML file and return the DisplayString.Name property of the given ID
#>
function Get-SCSMStringName($ID)
{
	return $ManagementPackXml.GetElementsByTagName("DisplayString") | ? ElementID -eq $ID | %{ $_.Name }
}

<#
	This function will query the XML file and return the ParentFolder property from the Folders object of the given ID
#>
function Get-SCSMFolderParent($ID)
{
	return $ManagementPackXml.GetElementsByTagName("Folder") | ? ID -EQ $ID | %{ $_.ParentFolder }
}

<#
	This function will query the XML file and return the Folder property from the FolderItems object of the given ID
#>
function Get-SCSMViewFolder($ID)
{
	return $ManagementPackXml.GetElementsByTagName("FolderItem") | ? ElementID -EQ $ID | %{ $_.Folder }
}

<#
	This function will create a new SCSM Folder ( Cireson folder only ) given the fodler name, and the parent ID where it should be saved. 
#>
function New-SCSMFolder($folderName, $ParentFolderID)
{
	# Create a GUID for the new folder	
	$folderGUID = ([GUID]::NewGuid().Guid).replace("-", '')
	# Create a GUID for the first category of the new folder
	$folderCategoryGUID = ([GUID]::NewGuid().Guid).replace("-", '')
	# Create a GUID for the second category of the new folder
	$folderCategory2GUID = ([GUID]::NewGuid().Guid).replace("-", '')
	# Set the description of the new folder
	$description = "Automatically created using script - For Details inquire with Jonatan.Bernal"
	
	# Create a new XML "Folder" Element
	$xmlElement_Folder = $ManagementPackXml.CreateElement("Folder")
	# Configure the element
	$xmlElement_Folder.SetAttribute('ID', "Folder.$folderGUID")
	$xmlElement_Folder.SetAttribute('Accessibility', "Public")
	$xmlElement_Folder.SetAttribute('ParentFolder', $ParentFolderID)
	# Append the child element on the structure.
	$ManagementPackXml.ManagementPack.Presentation.Folders.AppendChild($xmlElement_Folder) | Out-Null
	
	# Create a new XML "Category" Element
	$xmlElement_FolderCategory = $ManagementPackXml.CreateElement("Category")
	# Configure the element
	$xmlElement_FolderCategory.SetAttribute("ID", "Category.$folderCategoryGUID")
	$xmlElement_FolderCategory.SetAttribute("Target", "Folder.$folderGUID")
	$xmlElement_FolderCategory.SetAttribute("Value", "View!Cireson.View.Builder.Folder.Tasks")
	# Append the child element on the structure.
	$ManagementPackXml.ManagementPack.Categories.AppendChild($xmlElement_FolderCategory) | Out-Null
	
	# Create a new XML "Category" Element
	$xmlElement_2ndFolderCategory = $ManagementPackXml.CreateElement("Category")
	# Configure the element
	$xmlElement_2ndFolderCategory.SetAttribute("ID", "Category.$folderCategory2GUID")
	$xmlElement_2ndFolderCategory.SetAttribute("Target", "Folder.$folderGUID")
	$xmlElement_2ndFolderCategory.SetAttribute("Value", "View!Cireson.View.Builder.View.Tasks")
	# Append the child element on the structure.
	$ManagementPackXml.ManagementPack.Categories.AppendChild($xmlElement_2ndFolderCategory) | Out-Null
	
	# Create a new XML "ImageReference" Element
	$xmlElement_image = $ManagementPackXml.CreateElement("ImageReference")
	# Configure the element
	$xmlElement_image.SetAttribute("ElementID", "Folder.$folderGUID")
	$xmlElement_image.SetAttribute("ImageID", "Console!Microsoft.EnterpriseManagement.ServiceManager.UI.Console.Image.Folder")
	# Append the child element on the structure.
	$ManagementPackXml.ManagementPack.Presentation.ImageReferences.AppendChild($xmlElement_image) | Out-Null
	
	# Createa a new "DisplayString" Element
	$xmlElement_DisplayName = $ManagementPackXml.CreateElement("DisplayString")
	# Configure the element
	$xmlElement_DisplayName.SetAttribute('ElementID', "Folder.$folderGUID")
	# Createa a new "Name" Element
	$xmlElement_DisplayName_Name = $ManagementPackXml.CreateElement("Name")
	# Configure the element
	$xmlElement_DisplayName_Name.InnerText = $folderName
	# Append the child element to the new "DisplayString" element.
	$xmlElement_DisplayName.AppendChild($xmlElement_DisplayName_Name) | Out-Null
	# Create a new "Description" Element
	$xmlElement_DisplayName_Description = $ManagementPackXml.CreateElement("Description")
	# Configure the element
	$xmlElement_DisplayName_Description.InnerText = $description
	# Append the child element to the new "DisplayString" element.
	$xmlElement_DisplayName.AppendChild($xmlElement_DisplayName_Name) | Out-Null
	# Append the child element on the structure. 
	$ManagementPackXml.ManagementPack.LanguagePacks.LanguagePack | ? ID -eq "ENU" | %{ $_.DisplayStrings.AppendChild($xmlElement_DisplayName) } | Out-Null
	
	# Call the 'Save-Xml' function to save the changes
	Save-XML
	
	# return the ID of the newly created folder.
	return "Folder.$folderGUID"
} # endfunction

<#
	This function will create a copy of an SCSM View. It requires that a new name, the parent ID of where the view should be stored, and the original View ID be provided.
#>
function Copy-SCSMView($viewName, $parentID, $originalViewID)
{
	# Create a GUID for the new View
	$viewGUID = ([GUID]::NewGuid().Guid).replace("-", '')
	# Create a GUID for the new FolderItem
	$viewItemGUID = ([GUID]::NewGuid().Guid).replace("-", '')
	# Get the original view Object
	$xmlElement_originalView = $ManagementPackXml.ManagementPack.Presentation.Views.View | ? ID -match $originalViewID
	# Get the original view categories
	$xmlElement_originalViewCategories = $ManagementPackXml.ManagementPack.Categories.Category | ? Target -EQ $xmlElement_originalView.ID
	# Get the Original view image 
	$xmlElement_originalViewImage = $ManagementPackXml.ManagementPack.Presentation.ImageReferences.ImageReference | ? ElementID -EQ $xmlElement_originalView.ID
	# Clone the original view Element
	$xmlElement_View = $xmlElement_originalView.Clone()
	# Set the Clone ID to the new View ID
	$xmlElement_View.ID = "View.$viewGUID"
	# Append the child to "Views" 
	$ManagementPackXml.ManagementPack.Presentation.Views.AppendChild($xmlElement_View) | Out-Null
	
	foreach ($categoryObj in $xmlElement_originalViewCategories) # Loop through each of the found categories
	{
		# Create a GUID for the new category
		$viewCategoryGUID = ([GUID]::NewGuid().Guid).replace("-", '')
		# Create a "Category" Element
		$xmlElement_ViewCategory = $ManagementPackXml.CreateElement("Category")
		# Configure Element
		$xmlElement_ViewCategory.SetAttribute("ID", "Category.$viewCategoryGUID")
		$xmlElement_ViewCategory.SetAttribute("Target", "View.$viewGUID")
		$xmlElement_ViewCategory.SetAttribute("Value", $categoryObj.Value)
		# Append the child to "Categories"
		$ManagementPackXml.ManagementPack.Categories.AppendChild($xmlElement_ViewCategory) | Out-Null
	} # endloop
	
	# Create a "FolderItem" element
	$xmlElement_ViewItem = $ManagementPackXml.CreateElement("FolderItem")
	# Configure Element
	$xmlElement_ViewItem.SetAttribute('ElementID', "View.$viewGUID")
	$xmlElement_ViewItem.SetAttribute('ID', "FolderItem.$viewItemGUID")
	$xmlElement_ViewItem.SetAttribute('Folder', $parentID)
	# Append the child to "FolderITems" 
	$ManagementPackXml.ManagementPack.Presentation.FolderItems.AppendChild($xmlElement_ViewItem) | Out-Null
	
	# Create a new "ImageReference" Object
	$xmlElement_ViewImage = $ManagementPackXml.CreateElement("ImageReference")
	# Configure element
	$xmlElement_ViewImage.SetAttribute("ElementID", "View.$viewGUID")
	$xmlElement_ViewImage.SetAttribute("ImageID", $xmlElement_originalViewImage.ImageID)
	# Append the child to "ImageReferences"
	$ManagementPackXml.ManagementPack.Presentation.ImageReferences.AppendChild($xmlElement_ViewImage) | Out-Null
	
	# Createa a new "DisplayString" element
	$xmlElement_ViewDisplayName = $ManagementPackXml.CreateElement("DisplayString")
	# Configure element
	$xmlElement_ViewDisplayName.SetAttribute('ElementID', "View.$viewGUID")
	# Creata a new "Name" Element
	$xmlElement_ViewDisplayName_Name = $ManagementPackXml.CreateElement("Name")
	# Configure element
	$xmlElement_ViewDisplayName_Name.InnerText = $viewName
	# Append element to new DisplayString element
	$xmlElement_ViewDisplayName.AppendChild($xmlElement_ViewDisplayName_Name) | Out-Null
	# Create a new "Description" Element
	$xmlElement_ViewDisplayName_Description = $ManagementPackXml.CreateElement("Description")
	# Configure element
	$xmlElement_ViewDisplayName_Description.InnerText = "Automatically created using script - For Details inquire with Jonatan.Bernal"
	# Append element to new DisplayString element
	$xmlElement_ViewDisplayName.AppendChild($xmlElement_ViewDisplayName_Description) | Out-Null
	# Append element to "DisplayStrings"  
	$ManagementPackXml.ManagementPack.LanguagePacks.LanguagePack | ? ID -eq "ENU" | %{ $_.DisplayStrings.AppendChild($xmlElement_ViewDisplayName) } | Out-Null
	
	# Call the 'Save-Xml' function to save the changes
	Save-XML
	
	# return the ID of the newly created folder.  
	return "View.$viewGUID"
} # endfunction

<#
	This function will move a View to a different parent. It requires that the ID of the view and the ID of the new parent folder be provided.
#>
function Move-SCSMView($viewID, $newParentID)
{
	# Get the FolderItem XML element based on the provided ID
	$xmlViewItem = $ManagementPackXml.ManagementPack.Presentation.FolderItems.FolderItem | ? ElementID -EQ $viewID
	# Update the Folder property on the element
	$xmlViewItem.Folder = $newParentID
	
	if ($xmlViewItem.Folder -eq $newParentID) # Ensure the change was performed successfully.
	{
		# Call the 'Save-Xml' function to save the changes
		Save-XML
		return $true
	} # endif
	else # if the change could not be performed
	{
		return $false
	}  # endelse
} # endfunction

<#
	This function will move a Folder to a different parent. It requires that the folder ID and the new parent folder ID be provided.
#>
function Move-SCSMFolder($folderID, $newParentID)
{
	# Get the folder XML element
	$xmlFolderItem = $ManagementPackXml.ManagementPack.Presentation.Folders.Folder | ? ID -EQ $folderID
	# Upodate the parent folder property
	$xmlFolderItem.ParentFolder = $newParentID

	if ($xmlFolderItem.ParentFolder -eq $newParentID) # Ensure the change was performed successfully.

	{
		# Call the 'Save-Xml' function to save the changes
		Save-XML
		return $true
	} # endif
	else # if the change could not be performed
	{
		return $false
	} # endelse
} # endfunction

<#
	This function will remove an a View from the XML file. It requires that you provide the ID of the view that is to be removed.
	Changes performed:
		* Removes Categories Elements
		* Removes View Element
		* Removes the FolderItem Element 
		* Removes the ImageReference Element
		* Removes the DisplayString Element
#>
function Remove-SCSMView($viewID)
{
	# Get a list of all the Category elements
	$Categories = $ManagementPackXml.ManagementPack.Categories
	# Get a list of all the Views
	$Views = $ManagementPackXml.ManagementPack.Presentation.Views
	# Get a list of all the FolderItems
	$FolderItems = $ManagementPackXml.ManagementPack.Presentation.FolderItems
	# Get a list of all the ImageReferences
	$ImageReferences = $ManagementPackXml.ManagementPack.Presentation.ImageReferences
	# Get a list of all the DisplayString Elements
	$DisplayStrings = $ManagementPackXml.ManagementPack.LanguagePacks.LanguagePack | ? ID -eq "ENU" | %{ $_.DisplayStrings }
	
	foreach ($category in $Categories.Category) # Loop through each category
	{
		if ($category.Target -eq $viewID) # Find the one that contains the ViewID
		{
			# Remove the Category
			$Categories.RemoveChild($category)
		} # endif
	} # endloop
	
	foreach ($view in $Views.View) # Loop through each view
	{
		if ($view.ID -eq $viewId) # Find the one that contains the ViewID

		{
			# Remove the View
			$Views.RemoveChild($view)
		} # endif
	} # endloop
	
	foreach ($FolderItem in $FolderItems.FolderItem) # Loop through each FolderItem
	{
		if ($FolderItem.ElementID -eq $viewId) # Find the one that contains the ViewID
		{
			# Remove the FolderItem
			$FolderItems.RemoveChild($FolderItem)
		} # endif
	} # endloop
	
	foreach ($ImageReference in $ImageReferences.ImageReference) # Loop through each ImageReference
	{
		if ($ImageReference.ElementID -eq $viewId) # Find the one that contains the ViewID
		{
			# Remove the ImageReference
			$ImageReferences.RemoveChild($ImageReference)
		} # endif
	} # endloop
	
	foreach ($DisplayString in $DisplayStrings.DisplayString) # Loop through each DisplayString
	{
		if ($DisplayString.ElementID -eq $viewId) # Find the one that contains the ViewID
		{
			# Remove the DisplayString
			$DisplayStrings.RemoveChild($DisplayString)
		} # endif
	} # endloop
	
	# Call the 'Save-Xml' function to save the changes
	Save-XML
}

<#
	This function will remove an a Folder from the XML file. It requires that you provide the ID of the folder that is to be removed.
	Changes performed:
		* Removes Categories Elements
		* Removes Folder Element
		* Removes the FolderItem Element 
		* Removes the ImageReference Element
		* Removes the DisplayString Element
#>
function Remove-SCSMFolder($folderID)
{
	# Get a list of all the Category elements
	$Categories = $ManagementPackXml.ManagementPack.Categories
	# Get a list of all the Folders
	$Folders = $ManagementPackXml.ManagementPack.Presentation.Folders
	# Get a list of all the FolderItems
	$FolderItems = $ManagementPackXml.ManagementPack.Presentation.FolderItems
	# Get a list of all the ImageReferences
	$ImageReferences = $ManagementPackXml.ManagementPack.Presentation.ImageReferences
	# Get a list of all the DisplayString Elements
	$DisplayStrings = $ManagementPackXml.ManagementPack.LanguagePacks.LanguagePack | ? ID -eq "ENU" | %{ $_.DisplayStrings }
	
	foreach ($category in $Categories.Category) # Loop through each Folder
	{
		if ($category.Target -eq $folderID) # Find the one that contains the folderID
		{
			# Remove the Folder
			$Categories.RemoveChild($category)
		} # endif
	} # endloop
	
	foreach ($Folder in $Folders.Folder) # Loop through each Folder
	{ 
		if ($Folder.ID -eq $folderID) # Find the one that contains the folderID
		{
			# Remove the Folder
			$Folders.RemoveChild($Folder)
		} # endif
	} # endloop
	
	foreach ($FolderItem in $FolderItems.FolderItem) # Loop through each FolderItem
	{
		if ($FolderItem.Folder -eq $folderID) # Find the one that contains the folderID
		{
			if ($FolderItem.ElementID -like "View.*") # If the FolderItem is a view
			{
				# Call 'Remove-SCSMView' to remove the view from the XML file
				Remove-SCSMView -viewID $FolderItem.ElementID
			} # endif
			elseif ($FolderItem.ElementId -like "Folder.*") # IF the FolderItem is a folder
			{
				# Call 'Remove-SCSMFolder' to remove the folder and its subitems from the XML file
				Remove-SCSMFolder -folderID $FolderItem.ElementID
			} # endelse
		} # endif
	} # endloop
	
	foreach ($ImageReference in $ImageReferences.ImageReference) # Loop through each ImageReference
	{
		if ($ImageReference.ElementID -eq $folderID) # Find the one that contains the folderID
		{
			# Remove the ImageReference
			$ImageReferences.RemoveChild($ImageReference)
		} # endif
	} # endloop
	
	foreach ($DisplayString in $DisplayStrings.DisplayString) # Loop through each DisplayString
	{
		if ($DisplayString.ElementID -eq $folderID) # Find the one that contains the folderID
		{
			# Remove the DisplayString
			$DisplayStrings.RemoveChild($DisplayString)
		} # endif
	} # endloop
	
	# Call the 'Save-Xml' function to save the changes
	Save-XML
}

<#
	This function will rename the SCSM Object. This requires that the Object ID and the new name is provided.
#>
function Rename-SCSMObject($objectID, $newName)
{
	# Find the XML element equals to the provided ID
	$selectedItemNameElement = $ManagementPackXml.ManagementPack.LanguagePacks.LanguagePack | ? ID -eq "ENU" | %{ $_.DisplayStrings.DisplayString } | ? ElementID -EQ $objectID
	# Update the name property of the element
	$selectedItemNameElement.Name = $newName
	# Call the 'Save-Xml' function to save the changes
	Save-XML
} # endfunction