### This Scripts automates the Office 365 installation process
### It has to be run as and admin, if the ExecutionPolicy is Restricted
### run the scritp with as follows: powershell -ExecutionPolicy Bypass -File  "InstallO365.ps1"


# Enable verbose output
$VerbosePreference = "Continue"

try {
    # Create a temporary directory for the Office deployment
    Write-Host "Creating temporary directory..." -Verbose
    New-Item -Path 'C:\' -Name 'Temporary' -ItemType Directory -Force
    $Dest = 'C:\Temporary'
    New-Item -Path $Dest -Name 'Office' -ItemType Directory -Force
    Write-Host "Temporary directory created at $Dest" -Verbose
} 
catch {
    Write-Host "Error creating temporary directory: $_" -ForegroundColor Red
    exit 1
}

try {
    # Download the Office Deployment Tool (ODT)
    Write-Host "Downloading Office Deployment Tool..." -Verbose
    $OfficeLink = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=49117' -ErrorAction Stop
    $Links = $OfficeLink.Links
    foreach ($this in $Links) {
        if ($this.href.Contains('https://download.microsoft.com/download')) {
            $DownloadLink = $this.href
            Write-Host "Found download link: $DownloadLink" -Verbose
        }
    }
    if (-not $DownloadLink) {
        Write-Host "Download link not found." -ForegroundColor Red
        exit 1
    }
    
    # Download the Office deployment tool file
    Write-Host "Downloading Office Deployment Tool to $Dest\officedeploymenttool.exe" -Verbose
    Invoke-WebRequest -Uri $DownloadLink -UseBasicParsing -OutFile "$Dest\officedeploymenttool.exe" -ErrorAction Stop
    Write-Host "Downloaded Office Deployment Tool successfully" -Verbose
} 
catch {
    Write-Host "Error downloading Office Deployment Tool: $_" -ForegroundColor Red
    exit 1
}

try {
    # Create the installconfig.xml file for Office installation configuration
    Write-Host "Creating installconfig.xml for Office installation..." -Verbose
    New-Item -Path $Dest -ItemType File -Name 'installconfig.xml' -Force -ErrorAction Stop

    # Write the configuration XML to the file
    @'
    <Configuration>
      <Add OfficeClientEdition="64">
        <Product ID="O365BusinessRetail">
          <Language ID="en-us" />
        </Product>
      </Add>  
    </Configuration>
'@ | Set-Content -Path "$Dest\installconfig.xml" -Force -ErrorAction Stop
    Write-Host "installconfig.xml file created successfully" -Verbose
} 
catch {
    Write-Host "Error creating installconfig.xml: $_" -ForegroundColor Red
    exit 1
}

try {
    # Extract the setup file from the Office Deployment Tool
    Write-Host "Extracting Office Deployment Tool..." -Verbose
    Start-Process -FilePath "$Dest\officedeploymenttool.exe" -ArgumentList '/quiet /extract:c:\Temporary\Office' -Wait -ErrorAction Stop
    Write-Host "Office Deployment Tool extracted successfully" -Verbose
} 
catch {
    Write-Host "Error extracting Office Deployment Tool: $_" -ForegroundColor Red
    exit 1
}

try {
    # Install Office
    Write-Host "Starting Office installation..." -Verbose
    Start-Process -FilePath 'c:\Temporary\Office\setup.exe' -ArgumentList '/configure c:\Temporary\installconfig.xml' -Wait -ErrorAction Stop
    Write-Host "Office installed successfully" -Verbose
} 
catch {
    Write-Host "Error installing Office: $_" -ForegroundColor Red
    exit 1
}

try {
    # Remove the Temporary folder
    Write-Host "Cleaning up: Removing Temporary folder..." -Verbose
    Remove-Item -Path 'C:\Temporary' -Recurse -Force -ErrorAction Stop
    Write-Host "Temporary folder removed" -Verbose
} 
catch {
    Write-Host "Error removing Temporary folder: $_" -ForegroundColor Red
    exit 1
}
