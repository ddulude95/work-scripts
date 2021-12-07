# Author: Devin Dulude
# For York Hospital
#
# Info: This is a script I made using setuserfta by Christoph Kolbicz to set default programs on user profiles
#
#
# This runs once at startup and checks for the default apps of web browser, mail, pdf viewer, and word files. Then sets default apps accoridngly.
# With Paragon 20, we need edge as the default web browser, so we will be setting Edge as the default web browser. 
# We will allow Chrome, but IE will get changed to Edge. We also make sure
# .docx files are associated with Word, pdf viewer is Adobe, and mail client is Outlook as opposed to Win10 mail.
#
# This script is easily modifiable as I have stored everything in variables at the top. So, if the location for a file moves, you can change the directory for the variable.
# 

# this is just a 5 sec sleep for testing purposes
#$wsh = New-Object -ComObject Wscript.Shell
#$wsh.Popup("Starting Sleep")
#Start-Sleep -s 5
#$wsh.Popup("Running Script")



#setting up variables to the setuserfta cfg files
$existuser_edge = "fta_existuser_edge.txt" # sets .docx. to word mailto, to outlook, and http and https to Edge
$existuser_full = "fta_existuser_full.txt" # sets .pdf as adobe, .docx. to word mailto, to outlook, and http and https to Edge
$defaults_pdf = "fta_pdf.txt" # sets .pdf as adobe, docx to wait, mailto to outlook, and doesnt touch http and https
$defaults_misc = "fta_others.txt" # sets docx to word and mailto to outlook

# where the exe and cfg files live
$exePath = "\\yhdomainsvr1\netlogon\setuserfta\setuserfta.exe"
$cfgPath = "\\yhdomainsvr1\netlogon\setuserfta\"

# setting up paths / commands to be used
$ftaArgs = $exePath + " " + $cfgPath
$ftaGetBrowser = 'get | find "https"'
$ftaGetPdf = 'get | find "pdf"'
$ftaGetMail = 'get | find "mailto"'
$ftaGetDoc = 'get | find "docx"'

# the text file that we will write output to
$textFilePath = "$home\setuserfta.txt"

# setting these next two variables to 0 to act as booleans
$textFile = 0
$writeToFile = 0

# to be used with our switch statement 
$option = 0

# storing program IDs
$edgeLegacyPdfID = "AppXd4nrz8ff68srnhf9t5a8sbjyar1cr723"
$edgePdfID = "MSEdgePDF"
$edgeBrowserID = "MSEdgeHTM"
$edgeLegacyBrowserID = "AppX90nv6nhay5n6a98fnetv7tpk64pp35es"
$chromeBrowserID = "ChromeHTML"
$IESBrowserID = "IE.HTTPS"
$IEBrowserID = "IE.HTTP"
$wordID = "Word.Document.12"
$mailID = "Outlook.URL.mailto.15"

#checking to see if a text file exists
if (Test-Path $textFilePath -PathType Leaf)
{
    $textFile = 1
}

# running setuserfta.exe get to find out what the current profile is using
$findBrowser = cmd.exe /c "$exePath $ftaGetBrowser"
$findPdf = cmd.exe /c "$exePath $ftaGetPdf"
$findMail = cmd.exe /c "$exePath $ftaGetMail"
$findDoc = cmd.exe /c "$exePath $ftaGetDoc"

#---------- looking up what each program is set to as its default-------------#
# -------- FOR loop to separate the output of the get (it returns the extension and the prog ID) ------------ #
# trimming off any blank spaces in each of the split elements
$browserElements = $findBrowser.Split(",")
for ($i = 0; $i -lt $browserElements.Length; $i++)
{
    $browserElements[$i] = $browserElements[$i].Trim()
}

$pdfElements = $findPdf.Split(",")
for ($i = 0; $i -lt $pdfElements.Length; $i++)
{
    $pdfElements[$i] = $pdfElements[$i].Trim()
}

$mailElements = $findMail.Split(",")
for ($i = 0; $i -lt $mailElements.Length; $i++)
{
    $mailElements[$i] = $mailElements[$i].Trim()
}


# sometimes findDoc can be null because it's not set to anything. Quick check first before running this segment
if ($findDoc) {
    $docElements = $findDoc.Split(",")
    for ($i = 0; $i -lt $docElements.Length; $i++)
    {
        $docElements[$i] = $docElements[$i].Trim()
    }
}
#-----------------------------------------------------------#



#-----------------------------------------------------------#
# -------- checking for default apps and setting accordingly ------------ #
# need to check for both versions of Edge, old and new, they use two different program IDs

# this scenario covers a brand new profile on windows, should be edge set as PDF viewer and edge as browser by windows default
if ( ( ($edgePdfID -eq $PdfElements[1]) ) -and ( ($edgeBrowserID -eq $browserElements[1]) ) ) 
{
    $results = "Edge found as default PDF AND web browser"
    Write-Host $results
    $ftaArgs+= $defaults_pdf
    $writeToFile = $results + "`nOption 1. Setting pdf viewer as adobe, leaving Edge as browser, and setting mailto and docx extensions"
    $option = 1
}
# this scenario covers an exisitng user whose browser is IE but default pdf is edge (idk if this could even happen)
elseif ( ( ($edgePdfID -eq $PdfElements[1]) ) -and ( ($IESBrowserID -eq $browserElements[1]) -or ($IEBrowserID -eq $browserElements[1]) ) )
{
    $results = "Edge found as default PDF viewer and IE as default browser"
    Write-Host $results
    $ftaArgs+= $existuser_full
    $writeToFile = $results + "`nOption 2. Setting Adobe as default pdf viewer, Edge as browser, and mailto and docx"
    $option = 2
}
# this scenario checks for IE as the browser, and makes sure edge isnt pdf. most likely an existing user where Adobe is PDF and IE is browser. Wont touch .pdf in this case
# because we have some users who have a custom version of Adobe so we dont want to reset that if it's already set.
elseif ( ( ($IEBrowserID -eq $browserElements[1]) -or ($IESBrowserID -eq $browserElements[1]) ) -and ($edgePdfID -ne $PdfElements[1]) )
{
    $results = "IE found as default Internet Browser and edge NOT pdf viewer. pdf = " + $pdfElements[1]
    Write-Host $results
    $ftaArgs+= $existuser_edge
    $writeToFile = $results + "`nOption 3. Setting Edge as default browser, as well as mail and docx"
    $option = 3
}
# If edge is the pdf viewer and IE is not the browser (either firefox or chrome, because edge as pdf and edge as browser was already the top condition) 
# only set the pdf viewer, and also mailto and docx.
elseif ( ($pdfElements[1] -eq $edgePdfID) -and ( ($browserElements[1] -ne $IEBrowserID) -or ($browserElements[1] -ne $IESBrowserID) ) )
{
    $results = "Edge found as pdf viewer, and browser not IE. Browser = " + $browserElements[1]
    Write-Host $results
    $ftaArgs+= $defaults_pdf
    $writeToFile = $results + "`nOption 4. Setting Adobe as PDF viewer and Word as mailto and .docx associations"
    $option = 4
}

# finally, if everything is set how it should be, at least check for docx and mailto
elseif ( ($mailElements[1] -ne $mailID) -or ($docElements[1] -ne $wordID) )
{
    $results = "Adobe found as PDF viewer and Browser is NOT IE. Either docx or mailto set incorrectly."
    Write-Host $results
    $ftaArgs+= $defaults_misc
    $writeToFile = $results + "`nOption 5. Setting Outlook and Word as mailto and .docx associations"
    $option = 5
}
#-----------------------------------------------------------#




#-----------------------------------------------------------#
# -- switch statement to determine what we are going to do --- #
switch($option)
{
    0 {"Option 0. Browser is " + $browserElements[1] + " and Adobe is PDF. Doing nothing."}
    1 {"Option 1. Setting pdf viewer as adobe, leaving Edge as browser, and setting mailto and docx extensions"; cmd.exe /c $ftaArgs }
    2 {"Option 2. Setting Adobe as default pdf viewer, Edge as browser, and mailto and docx"; cmd.exe /c $ftaArgs }
    3 {"Option 3. Setting Edge as default browser, as well as mail and docx"; cmd.exe /c $ftaArgs }
    4 {"Option 4. Setting Adobe as PDF viewer and Word as mailto and .docx associations"; cmd.exe /c $ftaArgs }
    5 {"Option 5. Setting Outlook and Word as mailto and .docx associations"; cmd.exe /c $ftaArgs }
        
}
# ----------------------------------------------------- #
if ($writeToFile){

    if (-not $textFile)
    {
        New-Item $textFilePath
        Add-Content $textFilePath -Value (Get-Date)
        Add-Content $textFilePath -Value "$writeToFile `n"
    }
    else
    {
        Add-Content $textFilePath -Value (Get-Date)
        Add-Content $textFilePath -Value "$writeToFile `n"
    }
}

