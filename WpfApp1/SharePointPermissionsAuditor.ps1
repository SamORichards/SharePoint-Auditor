
#Connect to SharePoint Tenant Admin site. Input your tenant SharePoint domain here. 
param([string] $AdminCenterURL = "https://ocdtechbhop-admin.sharepoint.com",
	[string] $AdminEmail = "hnguyen@ocdtechbhop.onmicrosoft.com",
    [string] $FinalReport = "C:\temp\SitePermissionReport_ItemLevelScan.csv")
#Change your preferred directory here, if not you can check your local computer temporary directory. Open Report file, not Data file after execution. 
$CurrentDatey = Get-Date
$CurrentDate = $CurrentDatey.ToString('MM-dd-yyyy_hh-mm-ss')


#Check if GAC modules are installed due to conflicts of modules to run PnP commands
$GACUninstall = "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\"
$UninstallList = @() 
$UninstallList += Get-ChildItem -Path $GACUninstall -Recurse | Where-Object {$_.Title-contains "Microsoft.SharePoint." } | Select-Object FullName
if ($UninstallList.Count -eq 0) {
    Write-Host "Modules are compliant! Ready to execute script." -f Green
} elseif ($UninstallList.Count -gt 0) {
    foreach ($gacModule in $UninstallList) {
        try {
            remove-item $gacModule
            Write-Host "Removed module " + "$gacModule" -f Green
        }
        catch {
            Write-Host -f Red "Error deleting GAC file: " $_.Exception.Message
        }
    }
}


#Check if SharePoint Online PnP PowerShell module has been installed
Try {
    Write-host "Checking if SharePoint Online PnP PowerShell Module is Installed..." -f Yellow -NoNewline
    $SharePointPnPPowerShellOnline  = Get-Module -ListAvailable "SharePointPnPPowerShellOnline"
 
    If(!$SharePointPnPPowerShellOnline)
    {
        Write-host "No!" -f Green
 
        #Check if script is executed under elevated permissions - Run as Administrator
        If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
        {  
            Write-Host "Please Run this script in elevated mode (Run as Administrator)! " -NoNewline
            Read-Host "Press any key to continue"
            Exit
        }
 
        Write-host "Installing SharePoint Online PnP PowerShell Module..." -f Yellow -NoNewline
        Install-Module SharePointPnPPowerShellOnline -Force -Confirm:$False
        Write-host "Done!" -f Green
    }
    Else
    {
        Write-host "Yes!" -f Green
        Write-host "Updating SharePoint Online PnP PowerShell Module..." -f Yellow  -NoNewline
        Update-Module SharePointPnPPowerShellOnline
        Write-host "Done!" -f Green
    }
}
Catch{
    write-host "Error: $($_.Exception.Message)" -foregroundcolor red
}


#Check if SharePoint Online Exchage PowerShell module has been installed
Try {
    Write-host "Checking if SharePoint Online Exchange PowerShell Module is Installed..." -f Yellow -NoNewline
    $SharePointExchangePowerShellOnline  = Get-Module -ListAvailable "ExchangeOnlineManagement"
 
    If(!$SharePointExchangePowerShellOnline)
    {
        Write-host "No!" -f Green
 
        #Check if script is executed under elevated permissions - Run as Administrator
        If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
        {  
            Write-Host "Please Run this script in elevated mode (Run as Administrator)! " -NoNewline
            Read-Host "Press any key to continue"
            Exit
        }
 
        Write-host "Installing SharePoint Online Exchange PowerShell Module..." -f Yellow -NoNewline
        Install-Module -Name ExchangeOnlineManagement 
        Write-host "Done!" -f Green
    }
    Else
    {
        Write-host "Yes!" -f Green
        Write-host "Updating SharePoint Online Exchange PowerShell Module..." -f Yellow  -NoNewline
        Import-Module -Name ExchangeOnlineManagement 
        Write-host "Done!" -f Green
    }
}
Catch{
    write-host "Error: $($_.Exception.Message)" -foregroundcolor red
}

#Connecting to Online SharePoint services
Connect-ExchangeOnline -UserPrincipalName $AdminEmail -ShowProgress $true
Connect-PnPOnline -Url $AdminCenterURL -UseWebLogin
#Enumerate all sites of the tenant's SharePoint.
$hubSite = Get-PnPTenantSite $AdminCenterURL
$hubSiteId = $hubSite.HubSiteId
$URLarray=@()
$sites = Get-PnPTenantSite -Detailed
$sites | Select-Object url | % { 
  $s = Get-PnPTenantSite $_.url 
  if ($s.hubsiteid -eq $hubSiteId -and $s.Url -like "*/sites/*"){
    $URLarray += $s.url
  }
}


#Multi-Threading Jobs start:
for ($i = 0; $i -lt $URLarray.Length; $i++) {
    $url = $URLarray[$i]
    $filePath = ".\SitePermissionsData" + $i.ToString() + ".csv"
    Start-Job -ScriptBlock { 
    $ReportFile = $using:filePath
    $CurrentURL = $using:url
#Connecting to Online SharePoint services
Connect-PnPOnline -Url $using:AdminCenterURL -UseWebLogin

#Function to Get Permissions Applied on a particular Object, such as: Web, List, Folder or List Item
Function Get-PnPPermissions([Microsoft.SharePoint.Client.SecurableObject]$Object)
{
    #Determine the type of the object
    Switch($Object.TypedObject.ToString())
    {
        "Microsoft.SharePoint.Client.Web"  { $ObjectType = "Site" ; $ObjectURL = $Object.URL; $ObjectTitle = $Object.Title }
        "Microsoft.SharePoint.Client.ListItem"
        { 
            If($Object.FileSystemObjectType -eq "Folder")
            {
                $ObjectType = "Folder"
                #Get the URL of the Folder 
                $Folder = Get-PnPProperty -ClientObject $Object -Property Folder
                $ObjectTitle = $Object.Folder.Name
                $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''),$Object.Folder.ServerRelativeUrl)
                $SitePnP = Get-PnPTenantSite $ObjectURL -ErrorAction SilentlyContinue
            }
            Else #File or List Item
            {
                #Get the URL of the Object
                Get-PnPProperty -ClientObject $Object -Property File, ParentList
                If ($Null -ne $Object.File.Name)
                {
                    $ObjectType = "File"
                    $ObjectTitle = $Object.File.Name
                    $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''),$Object.File.ServerRelativeUrl)
                    $SitePnP = Get-PnPTenantSite $ObjectURL -ErrorAction SilentlyContinue
                }
                else
                {
                    $ObjectType = "List Item"
                    $ObjectTitle = $Object["Title"]
                    #Get the URL of the List Item
                    $DefaultDisplayFormUrl = Get-PnPProperty -ClientObject $Object.ParentList -Property DefaultDisplayFormUrl                     
                    $ObjectURL = $("{0}{1}?ID={2}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''), $DefaultDisplayFormUrl,$Object.ID)
                    $SitePnP = Get-PnPTenantSite $ObjectURL -ErrorAction SilentlyContinue
                }
            }
        }
        Default 
        { 
            $ObjectType = "List or Library"
            $ObjectTitle = $Object.Title
            #Get the URL of the List or Library
            $RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder     
            $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''), $RootFolder.ServerRelativeUrl)
            $SitePnP = Get-PnPTenantSite $ObjectURL -ErrorAction SilentlyContinue
        }
    }
   
    #Get permissions assigned to the object
    try {
        Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments
        write-host -f Green "Access Granted. Pulling information from site: " $Object.URL
    }
    catch {
        write-host -f Red "Access Denied. Storing URL."
    }
    
 
    #Check if Object has unique permissions
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments
     
    #Loop through each permission assigned and extract details
    $PermissionCollection = @()
    Foreach($RoleAssignment in $Object.RoleAssignments)
    { 
        #Get the Permission Levels assigned and Member
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
 
        #Get the Principal Type: User, SP Group, AD Group
        $PermissionType = $RoleAssignment.Member.PrincipalType
    
        #Get the Permission Levels assigned
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name
 
        #Remove Limited Access
        $PermissionLevels = ($PermissionLevels | Where-Object { $_ -ne "Limited Access"}) -join ","
 
        #Leave Principals with no Permissions
        If($PermissionLevels.Length -eq 0) {Continue}
 
        #Get SharePoint group members
        If($PermissionType -eq "SharePointGroup")
        {
            #Get Group Members
            $GroupMembers = Get-PnPGroupMembers -Identity $RoleAssignment.Member.LoginName
                 
            #Leave Empty Groups
            If($GroupMembers.count -eq 0){Continue}
            $GroupUsers = ($GroupMembers | Select-Object -ExpandProperty Title) -join ","

            # $SiteTitleName = $SitePnP.title
            # $UserArray =@()
       
            # if ($GroupUsers.Contains("$SiteTitleName")) {
            #     #$link_Type = ""
            #     if ($GroupUsers.contains("Members")) {
            #         $EvaluatedUser = Get-PnPGroup -Identity "$SiteTitleName" -Includes Users
            #         #$link_Type = "Members"
            #         #$GroupUsers += ": "
            #     }
            #     elseif ($GroupUsers.contains("Owners")) {
            #         $EvaluatedUser = Get-PnPGroup -Identity "$SiteTitleName" -Includes Users
            #         #$link_Type = "Owners"
            #         #$GroupUsers += ": "
            #     }
            #     elseif ($GroupUsers.contains("Vistors")) {
            #         $EvaluatedUser = Get-PnPGroup -Identity "$SiteTitleName" -Includes Users
            #         #$link_Type = "Subscribers"
            #         #$GroupUsers += ": "
            #     }
            #     elseif ($GroupUsers.contains("Aggregators")) {
            #         $EvaluatedUser = Get-PnPGroup-Identity "$SiteTitleName" -Includes Users
            #         #$link_Type = "Aggregators"
            #         #$GroupUsers += ": "
            #     } 
            #     elseif ($GroupUsers.contains("Event Subscribers")) {
            #         $EvaluatedUser = Get-PnPGroup -Identity "$SiteTitleName" -Includes Users
            #         #$link_Type = "EventSubscribers"
            #         #$GroupUsers += ": "
            #     }
            #     $UserArray += $EvaluatedUser
            #     #(Get-UnifiedGroup -Identity $SiteTitleName | Get-UnifiedGroupLinks -LinkType $link_Type -ErrorAction SilentlyContinue) 
            #     foreach ($User in $UserArray) {
            #         #Add the Data to Object
            #         $Permissions = New-Object PSObject
            #         $Permissions | Add-Member NoteProperty Object($ObjectType)
            #         $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            #         $Permissions | Add-Member NoteProperty URL($ObjectURL)
            #         $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            #         $Permissions | Add-Member NoteProperty Users($User)
            #         $Permissions | Add-Member NoteProperty Type($PermissionType)
            #         $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            #         $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
            #         $PermissionCollection += $Permissions  
            #     }
            # } else 
            #{
               #Add the Data to Object
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($GroupUsers)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
            $PermissionCollection += $Permissions 
            }
        #}
        Else
        {
            #Add the Data to Object
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($RoleAssignment.Member.Title)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
            $PermissionCollection += $Permissions
        }
    }
    #Export Permissions to CSV File
    $PermissionCollection | Export-CSV $ReportFile -NoTypeInformation -Append
}
   
#Function to get sharepoint online site permissions report
Function Get-PnPSitePermissionRpt()
{
    [cmdletbinding()]     
    Param  
    (    
        [Parameter(Mandatory=$false)] [String] $SiteURL, 
        [Parameter(Mandatory=$false)] [String] $ReportFile,         
        [Parameter(Mandatory=$false)] [switch] $Recursive,
        [Parameter(Mandatory=$false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory=$false)] [switch] $IncludeInheritedPermissions       
    )  
    Try {
        #Connect to the Site
        try {
            Write-host -f Yellow "Getting Site Collection Administrators..."
            Connect-PnPOnline -URL $CurrentURL -UseWebLogin
            #Get the Web
            $Web = Get-PnPWeb
            $CannotConnect = $false
            $SiteAdmins = Get-PnPSiteCollectionAdmin
        }
        catch {
            write-host -f Red "Error connecting due to restricted access. $_.Exception.Message."
            $CannotConnect = $true
        }
        if ($CannotConnect) {
            Write-host -f Yellow "Cannot connect to site: $input".
        #Add the Data to Object
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Object("Access Denied")
        $Permissions | Add-Member NoteProperty Title($Web.Title)
        $Permissions | Add-Member NoteProperty URL($Web.URL)
        $Permissions | Add-Member NoteProperty HasUniquePermissions("Access Denied")
        $Permissions | Add-Member NoteProperty Users("Access Denied")
        $Permissions | Add-Member NoteProperty Type("Access Denied")
        $Permissions | Add-Member NoteProperty Permissions("Access Denied")
        $Permissions | Add-Member NoteProperty GrantedThrough("Access Denied")
               
        #Export Permissions to CSV File
        $Permissions | Export-CSV $ReportFile -NoTypeInformation
   
        } else {
            Write-host -f Yellow "Getting Site Collection Administrators..."
        #Get Site Collection Administrators
        $SiteAdmins = Get-PnPSiteCollectionAdmin
         
        $SiteCollectionAdmins = ($SiteAdmins | Select-Object -ExpandProperty Title) -join ","
        #Add the Data to Object
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Object("Site Collection")
        $Permissions | Add-Member NoteProperty Title($Web.Title)
        $Permissions | Add-Member NoteProperty URL($Web.URL)
        $Permissions | Add-Member NoteProperty HasUniquePermissions("TRUE")
        $Permissions | Add-Member NoteProperty Users($SiteCollectionAdmins)
        $Permissions | Add-Member NoteProperty Type("Site Collection Administrators")
        $Permissions | Add-Member NoteProperty Permissions("Site Owner")
        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")

        #Export Permissions to CSV File
        $Permissions | Export-CSV $ReportFile -NoTypeInformation
        }
        
 
        #Function to Get Permissions of All List Items of a given List
        Function Get-PnPListItemsPermission([Microsoft.SharePoint.Client.List]$List)
        {
            Write-host -f Yellow "`t `t Getting Permissions of List Items in the List:"$List.Title
  
            #Get All Items from List in batches
            $ListItems = Get-PnPListItem -List $List -PageSize 500
  
            $ItemCounter = 0
            #Loop through each List item
            ForEach($ListItem in $ListItems)
            {
                #Get Objects with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                If($IncludeInheritedPermissions)
                {
                    Get-PnPPermissions -Object $ListItem
                }
                Else
                {
                    #Check if List Item has unique permissions
                    $HasUniquePermissions = Get-PnPProperty -ClientObject $ListItem -Property HasUniqueRoleAssignments
                    If($HasUniquePermissions -eq $True)
                    {
                        #Call the function to generate Permission report
                        Get-PnPPermissions -Object $ListItem
                    }
                }
                $ItemCounter++
                Write-Progress -PercentComplete ($ItemCounter / ($List.ItemCount) * 100) -Activity "Processing Items $ItemCounter of $($List.ItemCount)" -Status "Searching Unique Permissions in List Items of '$($List.Title)'"
            }
        }
 
        #Function to Get Permissions of all lists from the given web
        Function Get-PnPListPermission($Web)
        {
            #Get All Lists from the web
            $Lists = Get-PnPProperty -ClientObject $Web -Property Lists
   
            #Exclude system lists
            $ExcludedLists = @("Service Desk","Events","Generic List","Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
            "Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Images", "site collection images"
            ,"Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List","Site Assets","Preservation Hold Library",
            "Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Pages")
             
            $Counter = 0
            #Get all lists from the web   
            ForEach($List in $Lists)
            {
                #Exclude System Lists
                If($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title)
                {
                    $Counter++
                    Write-Progress -PercentComplete ($Counter / ($Lists.Count) * 100) -Activity "Exporting Permissions from List '$($List.Title)' in $($Web.URL)" -Status "Processing Lists $Counter of $($Lists.Count)"
 
                    #Get Item Level Permissions if 'ScanItemLevel' switch present
                    If($ScanItemLevel)
                    {
                        #Get List Items Permissions
                        Get-PnPListItemsPermission -List $List
                    }
 
                    #Get Lists with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-PnPPermissions -Object $List
                    }
                    Else
                    {
                        #Check if List has unique permissions
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments
                        If($HasUniquePermissions -eq $True)
                        {
                            #Call the function to check permissions
                            Get-PnPPermissions -Object $List
                        }
                    }
                }
            }
        }
    
   
        #Function to Get Webs's Permissions from given URL
        Function Get-PnPWebPermission($Web) 
        {
            #Call the function to Get permissions of the web
            Write-host -f Yellow "Getting Permissions of the Web: $($Web.URL)..." 
            Get-PnPPermissions -Object $Web
   
            #Get List Permissions
            Write-host -f Yellow "`t Getting Permissions of Lists and Libraries..."
            Get-PnPListPermission($Web)
 
            #Recursively get permissions from all sub-webs based on the "Recursive" Switch
            If($Recursive)
            {
                #Get Subwebs of the Web
                $Subwebs = Get-PnPProperty -ClientObject $Web -Property Webs
 
                #Iterate through each subsite in the current web
                Foreach ($Subweb in $web.Webs)
                {
                    #Get Webs with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-PnPWebPermission($Subweb)
                    }
                    Else
                    {
                        #Check if the Web has unique permissions
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $SubWeb -Property HasUniqueRoleAssignments
   
                        #Get the Web's Permissions
                        If($HasUniquePermissions -eq $true) 
                        { 
                            #Call the function recursively                            
                            Get-PnPWebPermission($Subweb)
                        }
                    }
                }
            }
        }
 
        #Call the function with RootWeb to get site collection permissions
        Get-PnPWebPermission $Web
   
        Write-host -f Green "`n*** Site Permission Report Generated Successfully! Appended Data to Final Report."
     }
    Catch {
        write-host -f Red "Error Generating Site Permission Report!" $_.Exception.Message
   }
}
   
#region ***Parameters***
    Get-PnPSitePermissionRpt -SiteURL $CurrentURL -ReportFile $ReportFile -Recursive -ScanItemLevel -IncludeInheritedPermissions
} 
}

Get-job | Wait-Job
#Exporting to final file
$Final = @()
for ($i = 0; $i -lt $URLarray.Length; $i++) {
    $DataPath = "C:\temp\SitePermissionsData" + $i.ToString() + ".csv"
    $Final += Import-Csv -path $DataPath 
}
$Final | Export-Csv -path $FinalReport -NotypeInformation -Force 

#endregion