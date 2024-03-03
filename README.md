# ExplorerBrowserCustomView
Display search results or other custom file set in IExplorerBrowser 

This is a port of [my VB6 project](https://www.vbforums.com/showthread.php?865485) that uses `IResultsFolder` to display search results in an `IExplorerBrowser` frame on your form:

Update (03 Mar 2024): .twinproj has been updated to use a more recent version of WinDevLib (formerly tbShellLib) due to errors in the package tB did not raise at the time this project was released.

![image](https://github.com/fafalone/ExplorerBrowserCustomView/assets/7834493/28aa8be9-948e-4509-9c0d-625b0e1323c3)

### About
Most of my shell projects thus far have focused on displaying locations in the shell, but what if you wanted to display the results of a search or some other list that involves files from all across the system? Turns out this is fairly straightforward to do in `IExplorerBrowser` using the `IResultsFolder` interface.

The demo project does this as a search, when the form comes up it first displays an empty ExplorerBrowser until you hit run. The startup routine then queries this blank display for its `IResultsFolder` interface, which represents the items in it-- none now, but the search will fill it.

Code:
```vba
'(...)
pEBrowse.Initialize Frame1.hWnd, rcEB, tFS
pEBrowse.SetOptions lFlag
pEBrowse.FillFromObject Nothing, EBF_NODROPTARGET
pEBrowse.GetCurrentView IID_IFolderView2, pfv
If (pfv Is Nothing) = False Then
'    'OPTIONAL
    'Customize which columns appear: fill uCol with however many columns (PROPERTYKEY's from mPKEY)
    ' you want, then be sure to change the second argument in SetColumns to the number of keys
    Dim uCol() As PROPERTYKEY
    ReDim uCol(2)
    Set pColMgr = pfv
    uCol(0) = PKEY_ItemNameDisplay
    uCol(1) = PKEY_ItemFolderPathDisplay
    uCol(2) = PKEY_Image_Dimensions
    pColMgr.SetColumns uCol(0), 3&
    pfv.GetFolder IID_IResultsFolder, pResFolder
Else
    Debug.Print "Error->No folderview"
End If
```

When you click run, it searches the given folder... you can modify this in any number of ways. The search algorithm here uses shell interfaces to enumerate and PathMatchSpec to compare. You can change search methods, change it to search multiple directories, etc. All that's important is that when you find a match, you add it to the `IResultsFolder`. The demo adds by `IShellItem`, but you can also add by pidl. The view is updated automatically as the search runs. The result is like the main picture up top.\
In the code, once the initialization routine has run, the IResultsFolder object is created, so all you have to do is call `.AddItem` or `.AddIDList`.

### Custom Columns
It's also possible to customize which columns you want to show.

In the demo project, there's optional code (commented out by default) in the initialization routine that shows how to make a list of PROPERTYKEY values to show as the column list.  


### Requirements
-Windows Vista or newer\
-Windows Development Library for twinBASIC
