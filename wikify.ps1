# Wiki generator script
# Input - path to documents
# Output - Wiki-fied text of Word docs and text files
#        - hierarchy of directories as pages
#        - embedded videos of MP4 files
#        - links to any other file


Import-Module VMware.VimAutomation.Core
Add-Type -AssemblyName "Microsoft.Office.Interop.Word"

# Set variables
$start_path = "C:\Temp\"
$wiki_server = "mhs-dal-wiki"
$vsphere_host = "vcenter.example.com"
$temp_wiki_path = "C:\Temp\wiki\"
$wiki_path = "/var/www/html/dokuwiki/data/pages/"
$wiki_image_path = "/var/www/html/dokuwiki/data/media/"
$vsphere_user = "domain\user"

# Main parsing function to create directory hierarchy pages
function make_directory_pages {
    Param([string]$path)
    
    Write-Host "Collecting all items in $path..."
    # Get all items in passed in directory
    $items = Get-ChildItem -Path $path 

    # Initialize current page
    $current_page = ""
    $current_page_name = (Get-Item -Path $path | Select -ExpandProperty BaseName).ToString()
    Write-Host "Working on page $current_page_name..."
    # Set page header
    $current_page += "====== " + $current_page_name + " ======`n`n"
    $current_page_file = ((($current_item.FullName).ToString() -Replace "\..{3}$","") -Replace '\W','').ToLower()

    # Loop through all the items and create a page or link for each one
   $items | ForEach-Object  {
        $current_item = Get-Item -Path ($_.FullName).ToString()

        # Set the page name
        $page_name = $current_item.BaseName.ToString()
        $page_file = ((($current_item.FullName).ToString() -Replace "\..{3}$","") -Replace '\W','').ToLower()

        # If this is another directory, recursively call the parsing function

        if (($current_item.Attributes).ToString() -eq "Directory") {
            $current_page += "[[" + $page_file + "|" + $page_name + "]]`n`n"
            make_directory_pages ($current_item.FullName).ToString()
        }
        # Convert Word document to wiki page
        elseif (($current_item.Name).ToString() -match '\.doc' -or ($current_item.Name).ToString() -match '\.pdf') {
            $current_page += "[[" + $page_file + "|" + $page_name + "]]`n`n"
            convert_word_doc ($current_item.FullName).ToString()
        }
        # Embed Video
        elseif (($current_item.Name).ToString() -match '\.mp4') {
            $current_page += "{{" + ($current_item.FullName).ToString() + "}}`n`n"
        }
        # Embed image 
        elseif (($current_item.Name).ToString() -match '\.gif' -or ($current_item.Name).ToString() -match '\.jpg' -or ($current_item.Name).ToString() -match '\.png') {
            $current_page += "{{" + ($current_item.FullName).ToString() + "}}`n`n"
        }
        # Convert text!
        elseif (($current_item.Name).ToString() -match '\.txt') {
            $current_page += "[[" + $page_file + "|" + $page_name + "]]`n`n"
            Copy-Item ($current_item.FullName.ToString()) ($temp_wiki_path + $page_file + ".txt")
        }
        # Create link for all other files
        else {
            $current_page += "[[" + ($current_item.FullName).ToString() + "|" + $page_name + "]]`n`n"
        }


    }
    # Save page as text file
    $current_page | Out-File -FilePath ($temp_wiki_path + $current_page_file + ".txt") -Encoding ascii

}

# Read word documents and convert them to dokuwiki text
function convert_word_doc {
    Param ([string]$path)
    
    Write-Host "Converting $path to HTML..."
    # Create Word object
    $wordDoc = New-Object -ComObject Word.Application

    # Open Word file
    $openDoc = $wordDoc.Documents.Open($path)

    # Save as HTML
    $openDoc.SaveAs(($temp_wiki_path + (((Get-Item $path | Select -ExpandProperty FullName).ToString() -replace "\..*","") -Replace '\W','').ToLower() + ".html"),[ref] [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatHTML)

    
    $openDoc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
    $wordDoc.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordDoc)
    Remove-Variable wordDoc

    # Lookup table for HTML to WIKI conversion
    $lookup_table = @{
        "(?ms)<h1.*?>" = '===== '
        '</h1>' = ' ====='
        '</h2>' = ' ===='
        "(?ms)<h2.*?>" = '==== '
        '</h3>' = ' ==='
        "(?ms)<h3.*?>" = '=== '
        "(?ms)<h4.*?>" = '== '
        '</h4>' = ' =='
        "(?ms)<br.*?>" = "\\ `n"
        "(?ms)<b\s+?.*?>" = '**'
        '</b>' = '**'
        "(?ms)<i .*?>" = '//'
        '</i>' = '//'
        "(?ms)<u .*?>" = '__'
        '</u>' = '__'
        '<b>' = '**'
        '<i>' = '//'
        '<u>' = '__'
        "(?ms)<sub.*?>" = '<sub>'
        '</sub>' = '</sub>'
        "(?ms)<sup.*?>" = '<sup>'
        '</sup>' = '</sup>'
        "<(?ms)p.*?>" = ''
        '</p>' = "\\ `n"
        "(?ms)<span.*?>" = ''
        '</span>' = ''
        "(?ms)<a.*?href=`"(?<link>.*?)`".*?>(?<description>.*?)<\/a>" = '[[${link}|${description}]]'
        "(?ms)<!--.*?-->" = ''
        "(?ms)<html.*?>" = ''
        '</html>' = ''
        "(?ms)<body.*?>" = ''
        '</body>' = ''
        "(?ms)<img.*?src=`"(?<imagefolder>.*?)/(?<imagefile>.*?)`".*?>" = '{{${imagefolder}${imagefile}}}'
        "(?ms)<!\[if.*?\]>" = ''
        '<!\[endif\]>' = ''
        "(?ms)<head.*</head>" = ''
        "(?ms)<div.*?>" = ''
        '</div>' = ''
        "(?ms)<v:.*?>" = ''
        "</v.*?>" = ''
        "(?ms)<table.*?>" = ''
        '</table>' = ''
        "(?ms)<td.*?>" = ''
        '</td>' = ''
        "(?ms)<tr.*?>" = ''
        '</tr>' = ''
        "(?ms)<o:.*?>" = ''
        "(?ms)</o.*?>" = ''
        '&nbsp;' = ''
        "(?ms)<a.+?name.*?>" = ''
        "</a>" = ''
        "(?ms)<p class=MsoTitle>(?<title>.*?)</p>" = '====== ${title} ======'
        "(?ms)file:///" = ''
        '&gt;' = '>'
        '&lt;' = '<'
        
    }

    # Begin HTML parsing!
    Write-Host "Opening HTML..."
    $html = Get-Content -Path ($temp_wiki_path + (((Get-Item $path | Select -ExpandProperty FullName).ToString() -replace "\..*","") -Replace '\W','').ToLower() + ".html") -Raw

    Write-Host "Converting HTML to wiki text..."

    $lookup_table.GetEnumerator() | ForEach-Object {
        if ($html -match $_.Key) {
            $html = $html -replace $_.Key, $_.Value
        }
    }

    $html += "Original document: [[$path]]"

    $html  -replace "`r","" -replace "·\s+",'  * ' -replace "\s+o\s+","`n`t`t* " -replace "§\s+","`t`t`t* " | Out-File -FilePath ($temp_wiki_path + (((Get-Item $path | Select -ExpandProperty FullName).ToString() -replace "\..*","") -Replace '\W','').ToLower() + ".txt") -Encoding ascii
    
    Remove-Variable html
}

function rename_images {
    Get-ChildItem -Path $temp_wiki_path -Include "*.jpg","*.png","*.gif" -Recurse | ForEach-Object {
        Move-Item -Path $_.FullName -Destination ($temp_wiki_path + ($_.DirectoryName -replace [regex]::Escape($temp_wiki_path),"") + $_.Name)
    }

}

# Call on base path
Write-Host "Initializing..."

# make_directory_pages $start_path

# Fix root directory
Copy-Item ($temp_wiki_path + ".txt") -Destination ($temp_wiki_path + "docs.txt")

# Move all images to root directory
# Write-Host "Fixing images..."
# rename_images


# Open connection to wiki server
# Gather authentication credentials
Write-Output "Please enter the following credentials: `n`n"

# Collect vSphere credentials
Write-Output "`n`nvSphere credentials:`n"
Write-Output "vSphere user: $vsphere_user"
$vsphere_pwd = Read-Host -Prompt "Enter the password for connecting to vSphere: " -AsSecureString

# Collect Linux credentials
Write-Output "`n`nRed Hat Linux credentials:`n"
Write-Output "Linux User: root"
$linux_user = "root"
$linux_pwd = Read-Host -Prompt "Enter the password for Linux: " -AsSecureString

# Create credential objects for all layers

$vsphere_creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $vsphere_user,$vsphere_pwd -ErrorAction Stop
$linux_creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $linux_user,$linux_pwd -ErrorAction Stop

Connect-VIServer -Server $vsphere_host -Credential $vsphere_creds -ErrorAction Stop

# Copy file to VM
Copy-VMGuestFile -Source ($temp_wiki_path + "*.txt") -Destination $wiki_path -VM $wiki_server -Confirm:$false -GuestUser $linux_creds.UserName -GuestPassword $linux_creds.Password -LocalToGuest -Force -Verbose -ToolsWaitSecs 120 -ErrorAction Stop

# Copy images to VM
Copy-VMGuestFile -Source ($temp_wiki_path + "*.jpg") -Destination $wiki_image_path -VM $wiki_server -Confirm:$false -GuestUser $linux_creds.UserName -GuestPassword $linux_creds.Password -LocalToGuest -Force -Verbose -ToolsWaitSecs 120 -ErrorAction Stop
Copy-VMGuestFile -Source ($temp_wiki_path + "*.png") -Destination $wiki_image_path -VM $wiki_server -Confirm:$false -GuestUser $linux_creds.UserName -GuestPassword $linux_creds.Password -LocalToGuest -Force -Verbose -ToolsWaitSecs 120 -ErrorAction Stop
Copy-VMGuestFile -Source ($temp_wiki_path + "*.gif") -Destination $wiki_image_path -VM $wiki_server -Confirm:$false -GuestUser $linux_creds.UserName -GuestPassword $linux_creds.Password -LocalToGuest -Force -Verbose -ToolsWaitSecs 120 -ErrorAction Stop


Write-Host "Fixing permissions..."
$permission_command = "chown -R apache:apache " + $wiki_path + "; chown -R apache:apache " + $wiki_image_path

Invoke-VMScript -VM $wiki_server -GuestCredential $linux_creds -ScriptText $permission_command -ErrorAction Stop

Write-Host "Done!"
