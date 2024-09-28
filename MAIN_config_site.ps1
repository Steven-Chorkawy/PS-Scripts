<#  START FUNCTION DEFINITIONS  #>
function MyConnect {
    param (
        [string]$Url
    )
    $getConn = Get-PnPWeb
    if ($getConn.Url -eq $Url) {
        Write-Host "Skipping connection" -ForegroundColor DarkMagenta
        return
    }
    Write-Host "`n`n" -ForegroundColor DarkMagenta
    Write-Host "Connecting to $($Url)" -ForegroundColor DarkMagenta
    Connect-PnPOnline -Url $Url -Interactive -ReturnConnection
}

function CreateLibraries {
    param (
        [string] $Input
    )
    Write-Host "CreateLibraries called."
}

function AddDefaultViews {
    param (
        [string] $input
    )
    Write-Host "AddDefaultViews called."
}

function ImportContentTypeTemplate {
    param (
        [string] $Input
    )
    Write-Host "ImportContentTypeTemplate called."
}

function PromptUserForScript {
    $scripts = @("Create Libraries", "Add Default Views", "Import Content Type Template")
    Write-Host "`nSelect Script to Run..."

    for ($i = 0; $i -lt $scripts.Count; $i++) {
        Write-Host "`t$($i+1). $($scripts[$i])"
    }

    return Read-Host "Enter Number"
}
<#  END FUNCTION DEFINITIONS    #>

<#  MAIN SCRIPT #>
Clear-Host

# MyConnect -Url "https://claringtonnet.sharepoint.com/sites/TemplateforCommitteeSites"

$selectedScript = PromptUserForScript

switch ($selectedScript) {
    1 { CreateLibraries }
    2 { AddDefaultViews }
    3 { ImportContentTypeTemplate }
    Default {
        Write-Host "'$($selectedScript)' is not a valid input." -ForegroundColor Red
    }
}

Write-Host "`n`nENDING MAIN_CONFIG_SITE"