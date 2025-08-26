# ============================================================================
# IEEE LaTeX Template Setup Script (Windows) - 2025 EDITION
# ============================================================================

param(
    [switch]$SkipPrompt
)

# Script configuration
$ErrorActionPreference = "Continue"
$WarningPreference = "Continue"

# Force UTF-8 for proper rendering
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

# Global success tracker
$Global:HasErrors = $false

# Get script directory for saving files
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) { $ScriptDir = Get-Location }

# Template configurations
$Templates = @{
    "1" = @{
        Name = "IEEE Conference Template"
        Description = "For conference proceedings and workshop papers"
        URL = "https://www.ieee.org/content/dam/ieee-org/ieee/web/org/conferences/conference-latex-template.zip"
        Folder = "IEEE-Conference"
    }
    "2" = @{
        Name = "IEEE Journal/Transaction Template"
        Description = "For journal articles and transactions papers"
        URL = "https://journals.ieeeauthorcenter.ieee.org/wp-content/uploads/sites/7/IEEE-Journal-Article-Template.zip"
        Folder = "IEEE-Journal"
    }
    "3" = @{
        Name = "IEEE Access Template"
        Description = "For IEEE Access journal submissions"
        URL = "https://ieeeaccess.ieee.org/wp-content/uploads/2024/04/ACCESS_latex_template_20240429.zip"
        Folder = "IEEE-Access"
    }
}

# ========== Helper Functions ==========
function Test-Command {
    param([Parameter(Mandatory = $true)][string]$Name)
    try { 
        Get-Command -Name $Name -ErrorAction Stop | Out-Null
        return $true 
    } catch { 
        return $false 
    }
}

function Write-ColoredText {
    param(
        [string]$Text,
        [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
        [string]$Color = "White"
    )
    if ($null -eq $Text) { $Text = "" }
    try { Write-Host $Text -ForegroundColor $Color } catch { Write-Host $Text }
}

function Write-Step { 
    param([string]$Message) 
    Write-ColoredText ("=== " + $Message + " ===") "Cyan" 
}

function Write-Success { 
    param([string]$Message) 
    Write-ColoredText ("[OK] " + $Message) "Green" 
}

function Write-Warn { 
    param([string]$Message) 
    Write-ColoredText ("[WARN] " + $Message) "Yellow" 
}

function Write-Err { 
    param([string]$Message) 
    $Global:HasErrors = $true
    Write-ColoredText ("[ERROR] " + $Message) "Red" 
}

function Safe-Download {
    param([string]$Uri, [string]$OutFile)
    try {
        Write-ColoredText ("Downloading: " + [System.IO.Path]::GetFileName($Uri)) "Gray"
        Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -TimeoutSec 300
        return (Test-Path -LiteralPath $OutFile)
    } catch {
        Write-Err ("Download failed: " + $_.Exception.Message)
        return $false
    }
}

function Test-Prerequisites {
    Write-Step "Checking Prerequisites"
    
    $allGood = $true
    
    # Check MiKTeX
    if (-not (Test-Command pdflatex)) {
        Write-Err "MiKTeX not found. Install from: https://miktex.org/download"
        $allGood = $false
    } else {
        Write-Success "MiKTeX is installed"
    }
    
    # Check Perl
    if (-not (Test-Command perl)) {
        Write-Err "Perl not found. Install Strawberry Perl from: https://strawberryperl.com/"
        $allGood = $false
    } else {
        try {
            $perlVersion = & perl --version 2>&1 | Select-String "This is perl"
            if ($perlVersion) {
                Write-Success ("Perl is installed: " + $perlVersion.ToString().Trim())
            } else {
                Write-Success "Perl is installed"
            }
        } catch {
            Write-Success "Perl is installed"
        }
    }
    
    return $allGood
}

function Update-MiKTeX {
    Write-Step "Updating MiKTeX Configuration"
    
    $updated = $false
    
    # Refresh file database
    if (Test-Command initexmf) {
        try {
            Write-ColoredText "Refreshing MiKTeX database..." "Yellow"
            & initexmf --update-fndb --quiet
            Write-Success "Database refreshed"
            $updated = $true
        } catch {
            Write-Warn ("Database refresh failed: " + $_.Exception.Message)
        }
    }
    
    # Ensure latexmk package
    if (Test-Command mpm) {
        try {
            Write-ColoredText "Installing/updating latexmk..." "Yellow"
            & mpm --update-db --quiet
            & mpm --install=latexmk --quiet
            Write-Success "latexmk package ensured"
            $updated = $true
        } catch {
            Write-Warn ("Package management failed: " + $_.Exception.Message)
        }
    }
    
    # Re-register script engines
    if (Test-Command initexmf) {
        try {
            Write-ColoredText "Re-registering script engines..." "Yellow"
            & initexmf --force --mklinks --quiet
            Write-Success "Script engines updated"
            $updated = $true
        } catch {
            Write-Warn ("Script engine registration failed: " + $_.Exception.Message)
        }
    }
    
    if (-not $updated) {
        Write-Warn "Could not update MiKTeX configuration"
    }
}

function Show-TemplateMenu {
    Write-ColoredText ""
    Write-ColoredText "Available IEEE LaTeX Templates:" "Cyan"
    Write-ColoredText ""
    
    foreach ($key in $Templates.Keys | Sort-Object) {
        $template = $Templates[$key]
        Write-ColoredText ("[" + $key + "] " + $template.Name) "White"
        Write-ColoredText ("    " + $template.Description) "Gray"
    }
    
    Write-ColoredText ""
    Write-ColoredText "[Q] Quit" "White"
    Write-ColoredText ""
}

function Get-TemplateChoice {
    while ($true) {
        Show-TemplateMenu
        $choice = Read-Host "Select template (1-3, Q to quit)"
        
        if ($choice -match "^[Qq]$") {
            Write-ColoredText "Setup cancelled by user." "Yellow"
            exit 0
        }
        
        if ($Templates.ContainsKey($choice)) {
            return $Templates[$choice]
        }
        
        Write-ColoredText "Invalid choice. Please select 1, 2, 3, or Q." "Red"
        Write-ColoredText ""
    }
}

function Setup-IEEETemplate {
    param([hashtable]$SelectedTemplate)
    
    Write-Step ("Setting Up " + $SelectedTemplate.Name)
    
    $projectPath = Join-Path $ScriptDir $SelectedTemplate.Folder
    $zipFile = Join-Path $ScriptDir "ieee-template.zip"
    $tempDir = Join-Path $ScriptDir "temp-extract"
    
    # Clean up any existing files
    @($projectPath, $zipFile, $tempDir) | ForEach-Object {
        if (Test-Path -LiteralPath $_) {
            Remove-Item -LiteralPath $_ -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    
    # Download template
    if (-not (Safe-Download -Uri $SelectedTemplate.URL -OutFile $zipFile)) {
        Write-Err "Failed to download IEEE template"
        return $null
    }
    
    try {
        # Extract template
        Write-ColoredText "Extracting template..." "Yellow"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        Expand-Archive -Path $zipFile -DestinationPath $tempDir -Force
        
        # Find source directory
        $sourceDir = $tempDir
        $subDirs = Get-ChildItem -Path $tempDir -Directory -ErrorAction SilentlyContinue
        if ($subDirs.Count -eq 1) {
            $sourceDir = $subDirs[0].FullName
        }
        
        # Set up project directory
        New-Item -ItemType Directory -Path $projectPath -Force | Out-Null
        Copy-Item -Path (Join-Path $sourceDir "*") -Destination $projectPath -Recurse -Force
        
        # Find main TeX file
        $mainTex = Get-ChildItem -Path $projectPath -Recurse -Filter "*.tex" |
                   Where-Object { $_.Name -match "(ieee|conference|main|bare|access)" } |
                   Select-Object -First 1
                   
        if (-not $mainTex) {
            $mainTex = Get-ChildItem -Path $projectPath -Recurse -Filter "*.tex" | 
                      Select-Object -First 1
        }
        
        if (-not $mainTex) {
            Write-Err "No TeX file found in template"
            return $null
        }
        
        # Clean up temporary files
        Remove-Item -LiteralPath $zipFile -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-Success ("Template ready: " + $SelectedTemplate.Folder)
        Write-Success ("Main file: " + $mainTex.Name)
        
        return @{
            ProjectPath = $projectPath
            MainFile = $mainTex.Name
            WorkDir = $mainTex.Directory.FullName
            TemplateName = $SelectedTemplate.Name
        }
        
    } catch {
        Write-Err ("Template setup failed: " + $_.Exception.Message)
        return $null
    }
}

function Compile-LaTeX {
    param([string]$MainFile, [string]$WorkDir)
    
    Write-Step "Compiling LaTeX Document"
    
    $currentDir = Get-Location
    try {
        Set-Location -Path $WorkDir
        Write-ColoredText ("Compiling: " + $MainFile) "Yellow"
        
        # Try latexmk first (best option)
        if ((Test-Command latexmk) -and (Test-Command perl)) {
            Write-ColoredText "Using latexmk (recommended)..." "Yellow"
            $output = & latexmk -pdf -synctex=1 -interaction=nonstopmode -file-line-error $MainFile 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                Write-Success "Document compiled successfully with latexmk"
                return $true
            } else {
                Write-Warn "latexmk failed, trying pdflatex..."
            }
        }
        
        # Fallback to pdflatex
        if (Test-Command pdflatex) {
            Write-ColoredText "Using pdflatex (multiple passes)..." "Yellow"
            
            # Run pdflatex twice for cross-references
            & pdflatex -synctex=1 -interaction=nonstopmode -file-line-error $MainFile | Out-Null
            & pdflatex -synctex=1 -interaction=nonstopmode -file-line-error $MainFile | Out-Null
            
            $pdfFile = $MainFile -replace "\.tex$", ".pdf"
            if (Test-Path -LiteralPath $pdfFile) {
                Write-Success "Document compiled successfully with pdflatex"
                return $true
            } else {
                Write-Err "pdflatex failed to produce PDF"
                return $false
            }
        }
        
        Write-Err "No LaTeX compiler available"
        return $false
        
    } catch {
        Write-Err ("Compilation failed: " + $_.Exception.Message)
        return $false
    } finally {
        Set-Location -Path $currentDir
    }
}

function Open-InVSCode {
    param([string]$ProjectPath)
    
    if (Test-Command code) {
        try {
            Write-ColoredText "Opening in VS Code..." "Yellow"
            Start-Process -FilePath "code" -ArgumentList ('"' + $ProjectPath + '"') -NoNewWindow
            Write-Success "Project opened in VS Code"
            return $true
        } catch {
            Write-Warn ("VS Code launch failed: " + $_.Exception.Message)
            return $false
        }
    } else {
        Write-ColoredText "VS Code not found. Install from: https://code.visualstudio.com/" "Gray"
        return $false
    }
}

# ========== MAIN EXECUTION ==========
Write-ColoredText @"
+====================================================================+
|                       IEEE LATEX SETUP 2025                        |
|                          BY HUGO MARKOFF                           |
| (Takes credit if it works but takes no responsibility if it fails) |
+====================================================================+
"@ "Cyan"

Write-ColoredText ""
Write-ColoredText ("Setup will create project in: " + $ScriptDir) "White"

# Check prerequisites
if (-not (Test-Prerequisites)) {
    Write-Err "Prerequisites not met. Please install required software and try again."
    Read-Host "Press Enter to exit"
    exit 1
}

# Update MiKTeX configuration
Update-MiKTeX

# Get template selection
$selectedTemplate = Get-TemplateChoice

# Set up IEEE template
$template = Setup-IEEETemplate -SelectedTemplate $selectedTemplate
if (-not $template) {
    Write-Err "Template setup failed"
    Read-Host "Press Enter to exit"
    exit 1
}

# Compile document
$compiled = Compile-LaTeX -MainFile $template.MainFile -WorkDir $template.WorkDir

# Open in VS Code
$opened = Open-InVSCode -ProjectPath $template.ProjectPath

# Summary
Write-ColoredText ""
Write-Step "Setup Complete"
Write-ColoredText ("Template: " + $template.TemplateName) "White"
Write-ColoredText ("Project location: " + $template.ProjectPath) "White"
Write-ColoredText ("Main TeX file: " + $template.MainFile) "White"

if ($compiled) {
    Write-Success "PDF compiled successfully!"
    Write-ColoredText "Your IEEE paper template is ready for editing!" "Green"
} else {
    Write-Warn "Manual compilation may be needed"
    Write-ColoredText ("Try: latexmk -pdf " + $template.MainFile + " (in project folder)") "Gray"
}

Write-ColoredText ""

# Auto-close if everything worked
if (-not $Global:HasErrors -and $compiled) {
    Write-Success "All operations completed successfully!"
    Write-ColoredText "Closing in 3 seconds..." "Gray"
    Start-Sleep -Seconds 3
    exit 0
} else {
    if ($Global:HasErrors) {
        Write-Warn "Some issues occurred. Review the output above."
    }
    Read-Host "Press Enter to close"
    exit 1
}