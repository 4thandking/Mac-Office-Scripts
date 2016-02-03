Sub fontThemeImport(control As IRibbonControl)
'the subroutine needs to have the same name as that called by the custom ribbon xml
'the typical proces sinvolves first making the .ppam, and then generting customUI.xml in the custom UI Editor on windows (with which you open the .ppam)

Dim fileAccessGranted As Boolean
Dim filePermissionCandidates

'this is a user-created file on the desktop, created by using the 'save command' on line 34
'assumes user has already edited the xml file'
fontScheme = "/Users/Chris/Desktop/PPTFontScheme.xml"
'prompt the user for a name to save the new theme as'
newScheme = InputBox("Enter a scheme name")

'fontlink is a symlink for:
'Library/Group\ Containers/UBF8T346G9.Office/User\ Content.localized/Themes.localized/Theme\ Fonts
'for some reason, this path needs to be found by using terminal and tabbing to get the space escapes and the .localized stuff
'the important note is that by default, ppt searches for files starting from /Users/Chris/Library/Containers/com.microsoft.Powerpoint/Data
'this folder contains several symlinks, which, if used, can allow write access to the whole system
'new symlinks are created in terminal with: ln -s /any/file/on/the/disk symlink-name

newSchemeDest = "fontlink/" & newScheme & ".xml"

'Create an array with file paths for which permissions are needed
'each file accessed only needs permission ONCE
filePermissionCandidates = Array(fontScheme)

'Request Access from User_
fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)

'returns true if access granted, false otherwise

'load the scheme
ActivePresentation.Designs(1).SlideMaster.Theme.ThemeFontScheme.Load fontScheme
'move the theme for the location for user font schemes'
ActivePresentation.Designs(1).SlideMaster.Theme.ThemeFontScheme.Save newSchemeDest

End Sub
