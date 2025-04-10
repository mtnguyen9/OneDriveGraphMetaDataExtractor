# Clean Microsoft Graph OneDrive File Extractor
# This script extracts file information from a OneDrive/SharePoint shared folder
# Usage: .\CleanGraphExtractor.ps1 -SharePointUrl "https://contoso-my.sharepoint.com/personal/user/Documents/Folder"

param (
    [Parameter(Mandatory=$true)]
    [string]$SharePointUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "OneDriveFileInfo_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Add System.Web for URL decoding
Add-Type -AssemblyName System.Web

# Ensure the proper modules are installed
function Ensure-ProperModules {
    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Sites",
        "Microsoft.Graph.Files"
    )
    
    $modulesInstalled = $true
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Required module $module is not installed." -ForegroundColor Yellow
            $modulesInstalled = $false
        }
    }
    
    if (-not $modulesInstalled) {
        Write-Host "`nWould you like to install the required modules now? (Y/N)" -ForegroundColor Cyan
        $installModules = Read-Host
        
        if ($installModules -eq "Y" -or $installModules -eq "y") {
            try {
                Write-Host "Installing Azure.Identity module (required dependency)..." -ForegroundColor Cyan
                Install-Module Azure.Identity -Scope CurrentUser -Force -AllowClobber
                
                Write-Host "Installing Microsoft.Graph modules..." -ForegroundColor Cyan
                Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
                
                # Verify installation
                Write-Host "`nVerifying installed modules:" -ForegroundColor Cyan
                Get-Module -Name Microsoft.Graph* -ListAvailable | 
                    Select-Object Name, Version | 
                    Format-Table -AutoSize
                
                # Import modules
                Import-Module Microsoft.Graph.Authentication
                Import-Module Microsoft.Graph.Sites
                Import-Module Microsoft.Graph.Files
                
                return $true
            }
            catch {
                Write-Host "Error installing modules: $_" -ForegroundColor Red
                return $false
            }
        }
        else {
            Write-Host "Module installation skipped. Script cannot continue without required modules." -ForegroundColor Red
            return $false
        }
    }
    
    # Import modules
    try {
        Import-Module Microsoft.Graph.Authentication
        Import-Module Microsoft.Graph.Sites
        Import-Module Microsoft.Graph.Files
        return $true
    }
    catch {
        Write-Host "Error importing modules: $_" -ForegroundColor Red
        return $false
    }
}

# Connect to Microsoft Graph with required scopes
function Connect-ToGraph {
    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        
        # Define required scopes
        $scopes = @(
            "Sites.Read.All",
            "Files.Read.All"
        )
        
        # Connect to Microsoft Graph
        Connect-MgGraph -Scopes $scopes
        
        # Check if connection was successful
        $context = Get-MgContext
        if ($null -eq $context) {
            throw "Failed to establish Microsoft Graph connection."
        }
        
        Write-Host "Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
        return $false
    }
}

# Parse SharePoint URL to extract components
function Parse-SharePointUrl {
    param (
        [string]$Url
    )
    
    $urlInfo = @{
        TenantDomain = $null
        SitePath = $null
        DocumentPath = $null
        IsPersonalSite = $false
    }
    
    # Clean URL (remove query parameters)
    if ($Url -match '([^?]*)') {
        $Url = $matches[1]
    }
    
    # Extract tenant domain
    if ($Url -match 'https://([^/]+)') {
        $urlInfo.TenantDomain = $matches[1]
    }
    
    # Handle personal sites
    if ($Url -match 'https://[^/]+/personal/([^/]+)') {
        $urlInfo.IsPersonalSite = $true
        $personPath = $matches[1]
        $urlInfo.SitePath = "/personal/$personPath"
    }
    # Handle team sites
    elseif ($Url -match 'https://[^/]+/sites/([^/]+)') {
        $sitePath = $matches[1]
        $urlInfo.SitePath = "/sites/$sitePath"
    }
    
    # Extract document path more precisely
    if ($Url -match '/Documents/(.*)') {
        # URL decode the document path
        $encodedPath = $matches[1].TrimEnd('/')
        $urlInfo.DocumentPath = [System.Web.HttpUtility]::UrlDecode($encodedPath)
    }
    elseif ($Url -match '/Documents$') {
        # If URL ends with /Documents, there's no further path
        $urlInfo.DocumentPath = ""
    }
    # Also check for direct path without "Documents" since this appears to be the case in your URL
    elseif ($Url -match '/personal/[^/]+/(.+)') {
        # This handles cases where there's a path after the personal site but no explicit "Documents" folder
        $encodedPath = $matches[1].TrimEnd('/')
        $urlInfo.DocumentPath = [System.Web.HttpUtility]::UrlDecode($encodedPath)
    }
    
    return $urlInfo
}

# Get site ID
function Get-SiteId {
    param (
        [string]$TenantDomain,
        [string]$SitePath
    )
    
    try {
        Write-Host "Attempting to retrieve site using SharePoint URL patterns..." -ForegroundColor Cyan
        $site = $null
        
        # Try personal site access
        if ($SitePath -match "/personal/") {
            Write-Host "Trying to access personal site with path: $SitePath" -ForegroundColor Cyan
            $siteIdPath = $TenantDomain + ":" + $SitePath
            $site = Get-MgSite -SiteId $siteIdPath -ErrorAction SilentlyContinue
        }
        
        # Try team site access
        if (-not $site -and $SitePath -match "/sites/") {
            Write-Host "Trying to access team site with path: $SitePath" -ForegroundColor Cyan
            $siteIdPath = $TenantDomain + ":" + $SitePath
            $site = Get-MgSite -SiteId $siteIdPath -ErrorAction SilentlyContinue
        }
        
        # Try tenant root site
        if (-not $site) {
            Write-Host "Trying to access root site..." -ForegroundColor Yellow
            $site = Get-MgSite -SiteId $TenantDomain -ErrorAction SilentlyContinue
        }
        
        # Try root keyword
        if (-not $site) {
            Write-Host "Trying to access using root identifier..." -ForegroundColor Yellow
            $site = Get-MgSite -SiteId "root" -ErrorAction SilentlyContinue
        }
        
        # Return if site found
        if ($site) {
            Write-Host "Found site: $($site.DisplayName) with URL: $($site.WebUrl)" -ForegroundColor Green
            return $site.Id
        }
        
        # Try team site with direct name
        if ($SitePath -match "/sites/([^/]+)") {
            $siteName = $matches[1]
            Write-Host "Trying to access site by direct name: $siteName" -ForegroundColor Yellow
            
            $siteIdPath = $TenantDomain + ":" + "/sites/$siteName"
            $site = Get-MgSite -SiteId $siteIdPath -ErrorAction SilentlyContinue
            
            if ($site) {
                Write-Host "Found site: $($site.DisplayName) with URL: $($site.WebUrl)" -ForegroundColor Green
                return $site.Id
            }
        }
        
        # Try alternative format for SharePoint Online
        if ($TenantDomain -match '([^\.]+)') {
            $tenantPrefix = $matches[1]
            
            if ($SitePath -match "/sites/([^/]+)") {
                $siteName = $matches[1]
                $siteIdFormat = "$tenantPrefix.sharepoint.com,sites,$siteName"
                Write-Host "Trying alternative site ID format: $siteIdFormat" -ForegroundColor Yellow
                
                $site = Get-MgSite -SiteId $siteIdFormat -ErrorAction SilentlyContinue
                
                if ($site) {
                    Write-Host "Found site using alternative format: $($site.DisplayName)" -ForegroundColor Green
                    return $site.Id
                }
            }
            elseif ($SitePath -match "/personal/([^/]+)") {
                $userPath = $matches[1]
                $siteIdFormat = "$tenantPrefix.sharepoint.com,personal,$userPath"
                Write-Host "Trying alternative personal site ID format: $siteIdFormat" -ForegroundColor Yellow
                
                $site = Get-MgSite -SiteId $siteIdFormat -ErrorAction SilentlyContinue
                
                if ($site) {
                    Write-Host "Found personal site using alternative format: $($site.DisplayName)" -ForegroundColor Green
                    return $site.Id
                }
            }
        }
        
        Write-Host "Could not retrieve the site using any of the standard approaches." -ForegroundColor Red
        return $null
    }
    catch {
        Write-Host "Error retrieving site ID: $_" -ForegroundColor Red
        return $null
    }
}

# Get drives in a site
function Get-SiteDrives {
    param (
        [string]$SiteId
    )
    
    try {
        $drives = Get-MgSiteDrive -SiteId $SiteId
        
        if ($drives -and $drives.Count -gt 0) {
            Write-Host "Found $($drives.Count) drives in site." -ForegroundColor Green
            return $drives
        }
        else {
            Write-Host "No drives found in site." -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Error retrieving site drives: $_" -ForegroundColor Red
        return $null
    }
}

# Get the document library drive
function Get-DocumentLibraryDrive {
    param (
        [array]$Drives
    )
    
    # Usually "Documents" is the main document library
    $documentDrive = $Drives | Where-Object { $_.Name -eq "Documents" }
    
    if ($documentDrive) {
        Write-Host "Found Documents library drive." -ForegroundColor Green
        return $documentDrive
    }
    
    # If not found by name, take the first drive (usually the document library)
    if ($Drives -and $Drives.Count -gt 0) {
        Write-Host "Using first available drive: $($Drives[0].Name)" -ForegroundColor Yellow
        return $Drives[0]
    }
    
    Write-Host "No suitable document library drive found." -ForegroundColor Red
    return $null
}

# Fixed recursive file information function
function Get-FileInfoRecursively {
    param (
        [string]$DriveId,
        [string]$ItemId,
        [string]$CurrentPath,
        [array]$FileList = @()
    )
    
    try {
        # Get children of the current item
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $ItemId
        
        # Add debug information
        Write-Host "Found $($children.Count) items in folder: $CurrentPath" -ForegroundColor DarkCyan
        
        foreach ($child in $children) {
            $childPath = if ($CurrentPath -eq "") { $child.Name } else { "$CurrentPath/$($child.Name)" }
            
            # FIXED LOGIC: If Folder property exists, treat as folder regardless of File property
            $isFolder = ($null -ne $child.Folder) -or 
                       ($null -ne $child.Package) -or 
                       ($child.MimeType -eq "application/x-ms-sharepoint") -or
                       ($child.MimeType -like "*folder*")
            
            # Debug output
            Write-Host "Item: $($child.Name), Path: $childPath" -ForegroundColor DarkGray
            Write-Host "  - File property: $(if ($child.File) { "Present" } else { "Absent" })" -ForegroundColor DarkGray
            Write-Host "  - Folder property: $(if ($child.Folder) { "Present" } else { "Absent" })" -ForegroundColor DarkGray
            Write-Host "  - MimeType: $($child.MimeType)" -ForegroundColor DarkGray
            Write-Host "  - IsFolder determination: $isFolder" -ForegroundColor DarkGray
            
            # If item is a file
            if (-not $isFolder) {
                Write-Host "Processing file: $childPath" -ForegroundColor DarkCyan
                
                # Create file info object
                $fileInfo = [PSCustomObject]@{
                    FileName = $child.Name
                    FileExtension = [System.IO.Path]::GetExtension($child.Name)
                    FilePath = $childPath
                    CreatedDateTime = $child.CreatedDateTime
                    ModifiedDateTime = $child.LastModifiedDateTime
                    CreatedBy = if ($child.CreatedBy.User.DisplayName) { $child.CreatedBy.User.DisplayName } else { "Unknown" }
                    ModifiedBy = if ($child.LastModifiedBy.User.DisplayName) { $child.LastModifiedBy.User.DisplayName } else { "Unknown" }
                    FileSize = [math]::Round(($child.Size / 1KB), 2)
                    ItemType = "File"
                }
                
                $FileList += $fileInfo
            }
            # If item is a folder, process recursively
            else {
                Write-Host "Entering folder: $childPath" -ForegroundColor Cyan
                
                # Optionally, add folder to the file list as well with a different ItemType
                $folderInfo = [PSCustomObject]@{
                    FileName = $child.Name
                    FileExtension = ""
                    FilePath = $childPath
                    CreatedDateTime = $child.CreatedDateTime
                    ModifiedDateTime = $child.LastModifiedDateTime
                    CreatedBy = if ($child.CreatedBy.User.DisplayName) { $child.CreatedBy.User.DisplayName } else { "Unknown" }
                    ModifiedBy = if ($child.LastModifiedBy.User.DisplayName) { $child.LastModifiedBy.User.DisplayName } else { "Unknown" }
                    FileSize = 0
                    ItemType = "Folder"
                }
                
                $FileList += $folderInfo
                
                # Process folder contents recursively
                $FileList = Get-FileInfoRecursively -DriveId $DriveId -ItemId $child.Id -CurrentPath $childPath -FileList $FileList
            }
        }
        
        return $FileList
    }
    catch {
        Write-Host "Error processing items: $_" -ForegroundColor Red
        return $FileList
    }
}

# Navigate to folder and get item
function Get-FolderItem {
    param (
        [string]$DriveId,
        [string]$Path
    )
    
    try {
        Write-Host "Attempting to access folder: $Path" -ForegroundColor Cyan
        $folderItem = $null
        
        # Try direct access first
        $pathFormats = @(
            "root:/$Path",
            "root:/Documents/$Path",
            "items/root:/$Path",
            "root" # fallback to root
        )
        
        foreach ($pathFormat in $pathFormats) {
            try {
                Write-Host "Trying to access with path format: $pathFormat" -ForegroundColor Yellow
                $folderItem = Get-MgDriveItem -DriveId $DriveId -DriveItemId $pathFormat -ErrorAction Stop
                Write-Host "Successfully accessed folder with path: $pathFormat" -ForegroundColor Green
                
                # If we accessed root but need to go deeper
                if ($pathFormat -eq "root" -and $Path) {
                    Write-Host "Accessed root, attempting to navigate to: $Path" -ForegroundColor Yellow
                    
                    # Navigate through path segments
                    $pathSegments = $Path.Split('/', [System.StringSplitOptions]::RemoveEmptyEntries)
                    Write-Host "Path segments to navigate: $($pathSegments -join ', ')" -ForegroundColor Yellow
                    $currentItem = $folderItem
                    
                    foreach ($segment in $pathSegments) {
                        if ($segment) {
                            Write-Host "Looking for segment: $segment" -ForegroundColor Cyan
                            $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $currentItem.Id
                            
                            # Print available items for debugging
                            Write-Host "Available items:" -ForegroundColor DarkGray
                            $children | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor DarkGray }
                            
                            # Try exact match
                            $nextItem = $children | Where-Object { $_.Name -eq $segment }
                            
                            # Try case-insensitive match
                            if (-not $nextItem) {
                                Write-Host "Exact match not found, trying case-insensitive match..." -ForegroundColor Yellow
                                $nextItem = $children | Where-Object { $_.Name -like $segment }
                            }
                            
                            # Try partial match
                            if (-not $nextItem) {
                                Write-Host "Case-insensitive match not found, trying partial match..." -ForegroundColor Yellow
                                $nextItem = $children | Where-Object { $_.Name -like "*$segment*" } | Select-Object -First 1
                            }
                            
                            if ($nextItem) {
                                $currentItem = $nextItem
                                Write-Host "Found matching item: $($nextItem.Name)" -ForegroundColor Green
                            }
                            else {
                                Write-Host "Could not find folder segment: $segment" -ForegroundColor Red
                                $currentItem = $null
                                break
                            }
                        }
                    }
                    
                    if ($currentItem) {
                        $folderItem = $currentItem
                        Write-Host "Successfully navigated to: $Path" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Failed to navigate to complete path" -ForegroundColor Red
                        continue
                    }
                }
                
                break
            }
            catch {
                Write-Host "Could not access with path format: $pathFormat - $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        return $folderItem
    }
    catch {
        Write-Host "Error accessing folder: $_" -ForegroundColor Red
        return $null
    }
}

# Main script execution
try {
    # Step 1: Ensure proper modules are installed
    $modulesReady = Ensure-ProperModules
    if (-not $modulesReady) {
        Write-Host "Cannot continue without the required modules. Exiting script." -ForegroundColor Red
        exit
    }
    
    # Step 2: Connect to Microsoft Graph
    $connected = Connect-ToGraph
    if (-not $connected) {
        Write-Host "Cannot continue without Microsoft Graph connection. Exiting script." -ForegroundColor Red
        exit
    }
    
    # Step 3: Parse the SharePoint URL
    Write-Host "`nParsing SharePoint URL: $SharePointUrl" -ForegroundColor Cyan
    $urlInfo = Parse-SharePointUrl -Url $SharePointUrl
    
    # If no folder path is specified, assume root
    if (-not $urlInfo.DocumentPath) {
        $urlInfo.DocumentPath = ""
    }
    
    # Add debugging to examine the parsed URL components
    Write-Host "URL components after parsing:" -ForegroundColor Magenta
    Write-Host "TenantDomain: [$($urlInfo.TenantDomain)]" -ForegroundColor Magenta
    Write-Host "SitePath: [$($urlInfo.SitePath)]" -ForegroundColor Magenta
    Write-Host "DocumentPath: [$($urlInfo.DocumentPath)]" -ForegroundColor Magenta
    Write-Host "IsPersonalSite: [$($urlInfo.IsPersonalSite)]" -ForegroundColor Magenta
    
    # Step 4: Get the site ID
    $siteId = Get-SiteId -TenantDomain $urlInfo.TenantDomain -SitePath $urlInfo.SitePath
    
    if (-not $siteId) {
        Write-Host "Cannot continue without a valid site ID. Exiting script." -ForegroundColor Red
        exit
    }
    
    # Step 5: Get the site drives
    $drives = Get-SiteDrives -SiteId $siteId
    
    if (-not $drives) {
        Write-Host "Cannot continue without access to site drives. Exiting script." -ForegroundColor Red
        exit
    }
    
    # Step 6: Get the document library drive
    $documentDrive = Get-DocumentLibraryDrive -Drives $drives
    
    if (-not $documentDrive) {
        Write-Host "Cannot continue without access to a document library drive. Exiting script." -ForegroundColor Red
        exit
    }
    
    # Step 7: Get file information
    Write-Host "`nRetrieving file information for path: $($urlInfo.DocumentPath)" -ForegroundColor Green
    
    # Get the folder item
    $folderItem = Get-FolderItem -DriveId $documentDrive.Id -Path $urlInfo.DocumentPath
    
    if ($folderItem) {
        Write-Host "Starting recursive file scan from: $($folderItem.Name)" -ForegroundColor Green
        $fileList = Get-FileInfoRecursively -DriveId $documentDrive.Id -ItemId $folderItem.Id -CurrentPath $urlInfo.DocumentPath
    }
    else {
        Write-Host "Could not access the specified folder path. Trying root folder." -ForegroundColor Yellow
        $fileList = Get-FileInfoRecursively -DriveId $documentDrive.Id -ItemId "root" -CurrentPath ""
    }
    
    # Step 8: Export the results
    if ($fileList -and $fileList.Count -gt 0) {
        $fileList | Export-Csv -Path $OutputFile -NoTypeInformation
        Write-Host "`nSuccessfully exported $($fileList.Count) files to: $OutputFile" -ForegroundColor Green
    }
    else {
        Write-Host "`nNo files were found or there was an error collecting file information." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
finally {
    # Disconnect from Microsoft Graph
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Cyan
    }
    catch {
        # Do nothing
    }
}