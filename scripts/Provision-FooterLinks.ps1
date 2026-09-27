<#
.SYNOPSIS
  Provisions the Footer Links list used by SPO-Hub Site Footer Application Customizer.

.EXAMPLE
  Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/intranet" -Interactive
  .\scripts\Provision-FooterLinks.ps1 -AddSampleItems
#>

[CmdletBinding()]
param(
  [string]$ListTitle = "Footer Links",
  [string]$ListUrl = "Lists/FooterLinks",
  [switch]$AddSampleItems
)

$ErrorActionPreference = "Stop"
Write-Host "Provisioning list '$ListTitle'..." -ForegroundColor Cyan

$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
  $list = New-PnPList -Title $ListTitle -Template GenericList -Url $ListUrl
  Write-Host "  Created list: $ListTitle" -ForegroundColor Green
} else {
  Write-Host "  List already exists: $ListTitle" -ForegroundColor Yellow
}

function Ensure-Field {
  param([string]$InternalName, [string]$DisplayName, [scriptblock]$Create)
  $existing = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
  if (-not $existing) {
    & $Create
    Write-Host "  Added field: $DisplayName ($InternalName)" -ForegroundColor Green
  } else {
    Write-Host "  Field exists: $DisplayName" -ForegroundColor DarkGray
  }
}

Ensure-Field -InternalName "Link" -DisplayName "Link" -Create {
  Add-PnPField -List $ListTitle -Type URL -InternalName "Link" -DisplayName "Link" -Required
}
Ensure-Field -InternalName "Url" -DisplayName "Url" -Create {
  Add-PnPField -List $ListTitle -Type Text -InternalName "Url" -DisplayName "Url" -AddToDefaultView
}
Ensure-Field -InternalName "Category" -DisplayName "Category" -Create {
  Add-PnPField -List $ListTitle -Type Choice -InternalName "Category" -DisplayName "Category" `
    -Choices "Company","Legal","Support","Resources","Social" -AddToDefaultView
  $f = Get-PnPField -List $ListTitle -Identity "Category"
  $f.FillInChoice = $true
  $f.Update()
  Invoke-PnPQuery
}
Ensure-Field -InternalName "DisplayOrder" -DisplayName "Display Order" -Create {
  Add-PnPField -List $ListTitle -Type Number -InternalName "DisplayOrder" -DisplayName "Display Order" -AddToDefaultView
  Set-PnPField -List $ListTitle -Identity "DisplayOrder" -Values @{ DefaultValue = "100" }
}
Ensure-Field -InternalName "IsActive" -DisplayName "Is Active" -Create {
  Add-PnPField -List $ListTitle -Type Boolean -InternalName "IsActive" -DisplayName "Is Active" -AddToDefaultView
  Set-PnPField -List $ListTitle -Identity "IsActive" -Values @{ DefaultValue = "1" }
}
Ensure-Field -InternalName "OpenInNewTab" -DisplayName "Open in new tab" -Create {
  Add-PnPField -List $ListTitle -Type Boolean -InternalName "OpenInNewTab" -DisplayName "Open in new tab"
  Set-PnPField -List $ListTitle -Identity "OpenInNewTab" -Values @{ DefaultValue = "1" }
}
Ensure-Field -InternalName "IconName" -DisplayName "Icon Name" -Create {
  Add-PnPField -List $ListTitle -Type Text -InternalName "IconName" -DisplayName "Icon Name"
}
Ensure-Field -InternalName "Description" -DisplayName "Description" -Create {
  Add-PnPField -List $ListTitle -Type Note -InternalName "Description" -DisplayName "Description"
}

$activeView = Get-PnPView -List $ListTitle -Identity "Active Links" -ErrorAction SilentlyContinue
if (-not $activeView) {
  Add-PnPView -List $ListTitle -Title "Active Links" -Fields "Title","Link","Category","DisplayOrder" `
    -Query "<Where><Eq><FieldRef Name='IsActive'/><Value Type='Boolean'>1</Value></Eq></Where><OrderBy><FieldRef Name='DisplayOrder' Ascending='TRUE'/></OrderBy>"
  Write-Host "  Created view: Active Links" -ForegroundColor Green
}

if ($AddSampleItems) {
  $samples = @(
    @{ Title = "About us"; Url = "/SitePages/About.aspx"; Category = "Company"; Order = 10; NewTab = $false }
    @{ Title = "Careers"; Url = "https://careers.contoso.com"; Category = "Company"; Order = 20; NewTab = $true }
    @{ Title = "Privacy"; Url = "/SitePages/Privacy.aspx"; Category = "Legal"; Order = 10; NewTab = $false }
    @{ Title = "Terms of use"; Url = "/SitePages/Terms.aspx"; Category = "Legal"; Order = 20; NewTab = $false }
    @{ Title = "Help desk"; Url = "https://support.contoso.com"; Category = "Support"; Order = 10; NewTab = $true }
    @{ Title = "Training"; Url = "/SitePages/Training.aspx"; Category = "Resources"; Order = 10; NewTab = $false }
  )
  foreach ($s in $samples) {
    $exists = Get-PnPListItem -List $ListTitle -PageSize 1 -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($s.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if (-not $exists) {
      Add-PnPListItem -List $ListTitle -Values @{
        Title = $s.Title
        Link = "$($s.Url), $($s.Title)"
        Category = $s.Category
        DisplayOrder = $s.Order
        IsActive = $true
        OpenInNewTab = $s.NewTab
      } | Out-Null
      Write-Host "  Sample item: $($s.Title)" -ForegroundColor Green
    }
  }
}

Write-Host "Done. Point FooterApplicationCustomizer listTitle to '$ListTitle'." -ForegroundColor Cyan
