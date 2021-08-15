 <#
.SYNOPSIS
	    Active Directory Powershell GUI
.NOTES
	Author		: Hayden Trail  
    email		: hayden@tailoredit.co.nz
    Company		: Tailored IT Solutions
    File Name	: powershellGUI.ps1
    
.VERSION CONTROL
    see https://github.com/haydentrail/ADPowershellGUI

.TODO
    - plugins for new tabs
    - pin tabs to retain
    - create updater script
 #>
 
$app = @{
    title="Active Directory Powershell GUI"
    version="0.1.3"
    repo="https://github.com/haydentrail/ADPowershellGUI/"
    repoRaw="https://raw.githubusercontent.com/haydentrail/ADPowershellGUI/"
    versionFile="main/version.info"
    readme="main/README.md"
    updater="main/updater.ps1"
}
function global:loadAssembly($assembly){
    try{[Reflection.Assembly]::Load($assembly) | Out-Null}
    catch{
        write-log "ERROR" "$assembly is required for this script. $($_.Exception.Message)"
        write-host "Press any key to exit";$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | out-null
        exit -1
    }
}
function loadModule($module,$commentError,$alternative,$commentAlways,$required){
    write-log "INFO" "Querying module $module"
    try{if (Get-Module $module){write-log "INFO" "$module is already imported"; return $true}}
    catch{}
    write-log "INFO" "Importing module $module"
    try{
        Import-Module $module -ErrorAction Stop
        if($commentAlways){write-log "WARN" $commentAlways}
        $true
    }catch{
        write-log "ERROR" "$($_.Exception.Message)"
        if($alternative){
            $global:adModuleDllLoaded = $true
            $alternative.keys | %{ loadModule $_ -commentAlways $alternative.$_ -required $required}
        }else{
            if($commentError){write-log "WARN" $commentError}
            if($required -ne $false){
                write-host "Press any key to exit";$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | out-null
                exit -1
            }
            $false
        }
    }
}
function write-log($level, $msg, $logonly=$false, $screenOnly=$false,$onlyDebug=$false){
    if($onlyDebug -eq $true -and $DebugLogging -eq $true){return}
    $now = get-date -Format "dd/MM/yyyy HH:mm:ss"
    if(!$logfile){$logfile="generic-log.log"}
    if($screenOnly -eq $false){Add-Content $logfile -Value "$now, $pid, $level, $msg"}
    if($logonly -eq $true){return}
    try{
        Switch ($level.tolower()){
            "warn"{write-host $msg -foregroundcolor "magenta"}
            "error"{write-host $msg -foregroundcolor "red"}
            default{write-host $msg}
        }
    }catch{}
}
function hashToString($hash,$asParameterString){
    if($asParameterString -eq $true){$hashstr = ""}else{$hashstr = "@{"}
    foreach ($key in $hash.keys) { 
        $v = $hash[$key]
        if($v -and ($v.GetType() -eq [pscredential])){$v = $($v.userName)}
        if($asParameterString -eq $true){$hashstr += "-$key `"$v`" "}
        else{if ($key -match "\s") { $hashstr += "`"$key`"" + "=" + "`"$v`"" + ";" } else { $hashstr += $key + "=" + "`"$v`"" + ";" } }
    }
    if($asParameterString -eq $false){$hashstr += "}"}
    return $hashstr 
}
# exportToExcel - tries to use the function export-excel, if the file is open or it cant write for some reason it will prompt to retry
function global:exportToExcel($objectsToExport,$excelFile,$sheetName,$format){#,$deleteSheetFirst){
    try{
        $params = @{AutoSize=$true;ClearSheet=$true;WorkSheetname=$sheetName;TableName=$sheetName;TableStyle="Medium6"}  
        write-log "INFO" "Calling: Export-Excel $excelFile $(hashToString $params $true)"
        [System.Windows.Forms.Application]::DoEvents() 
        $objectsToExport | Export-Excel $excelFile @params
        if($format.count){formatExcel $excelFile @($format+@{name=$sheetName})} #format needs to be passed as array of hashtables
        $true
    } catch {
        write-log "ERROR" $_.Exception.Message
        $_.Exception.Message
    }
}
# openExcelFile - tries to use the function Open-ExcelPackage, if the file is open or it cant write for some reason it will prompt to retry
function global:openExcelFile($excelFile){
    try{
        write-log "INFO" "Attempting to open to $excelFile"
        [System.Windows.Forms.Application]::DoEvents() 
        Open-ExcelPackage $excelFile
    } catch {
        write-log "ERROR" $_.Exception.Message
        $_.Exception.Message
    }    
}
# formatExcel - Formats sheets in excel file for headers, etc. (Requires ImportExcel module)
function global:formatExcel($xlfile,$xlWorkSheets){
    if(!$xlWorkSheets.count){write-log "WARN" "No sheet information was passed, no formatting will be done"; return}
    if(Test-Path -Path $xlfile -PathType Leaf){
        write-log "INFO" "Preparing $xlfile for formatting"
        $excel = openExcelFile $xlfile
        if([string]$excel.getType() -eq "string"){return}

        $psheet = @{FontName="verdana";FontSize=8;AutoSize=$true;}
        $ptitle = @{FontName="impact";FontSize=24;FontColor="red"}
        $psubTitle = @{FontName="calibri";FontSize=16;FontColor=[System.Drawing.ColorTranslator]::FromHtml('#203764')}
        $pheader = @{Bold=$true;BackgroundColor="black";FontColor="white";Underline=$true;AutoSize=$true}

        if($xlWorkSheets.count){
            $xlWorkSheets.ForEach({
                [System.Windows.Forms.Application]::DoEvents() 
                if($_.ContainsKey("name") -and $excel.Workbook.Worksheets[$_.name]){
                    write-log "INFO" "formatting worksheet $($_.name)"
                    try{
                        $ws = $excel.Workbook.Worksheets[$_.name]
                        Set-Format -Range $ws.dimension.address -Worksheet $ws @psheet
                        if($_.ContainsKey("type") -and $_.type -eq "ERROR") {
                            $psubTitle.FontColor = "red"
                            $ws.TabColor = 'red'
                        }
                        $headRow = 1
                        if($_.subtitle){
                            if($ws.CELLS["A1"].value -eq $_.subtitle -or $ws.CELLS["A2"].value -eq $_.subtitle){
                                write-log "INFO" "The worksheet $($_.name) already has subtitle set"
                                $containsData = $ws.dimension.rows -gt 2
                            }else{
                                write-log "INFO" "Inserting a row at A1 and setting subtitle to $($_.subtitle)"
                                $ws.InsertRow(1, 1)
                                $ws.CELLS["A1"].value = $_.subtitle
                            }   
                            if($ws.CELLS["A1"].value -eq $_.subtitle){Set-Format -Range "A1" -Worksheet $ws @psubTitle}else{Set-Format -Range "A2" -Worksheet $ws @psubTitle}
                            $headRow ++                 
                        }
                        if($_.title){
                            if($ws.CELLS["A1"].value -eq $_.title){
                                write-log "INFO" "The worksheet $($_.name) already has title set"
                            }else{
                                write-log "INFO" "Inserting a row at A1 and setting title to $($_.title)"
                                $ws.InsertRow(1, 1)
                                $ws.CELLS["A1"].value = $_.title
                            }
                            Set-Format -Range "A1" -Worksheet $ws @ptitle 
                            $headRow ++                 
                        }
                        if($ws.dimension -and $ws.dimension.rows -gt $headRow){
                            write-log "INFO" "Formatting the data with fixed header"
                            $endColumn = $ws.Dimension.End.address.Substring(0,$ws.Dimension.End.address.IndexOf([String]$ws.dimension.end.row))
                            #Set-Format -Range "A$($headRow):$($endColumn)$($headRow)" -Worksheet $ws @pheader
                            #$ws.Cells["A$($headRow):$($endColumn)$($headRow)"].AutoFilter = $true
                            $ws.View.FreezePanes($headRow+1,1)
                        }else{
                            write-log "WARN" "$($ws.name) contains no data"
                        }  
                    }catch{write-log "ERROR" $($_.Exception.Message)}            
                }else{
                    write-log "WARN" "$xlfile does not contain a worksheet as specified in '$(hashToString $_)'"
                }               
            })
        }else{
            write-log "WARN" "No worksheet information was passed in '$(hashToString $_)'"
        }
        Close-ExcelPackage $excel   
    }else{
        write-log "WARN" "Couldn't find $xlfile for formatting"
    }    
}
function global:addTab($tab,$scriptOnly=$false){
    write-log "INFO" "Adding $($tab.title) to gui."
    $html = @{
        tab="<button onclick=`"openTab(event, 'nav-$($Tab.id)','-main')`" class=`"nav-link nav-link-main`" id=`"nav-$($Tab.id)-tab`" data-bs-toggle=`"tab`" data-bs-target=`"#nav-$($Tab.id)`" type=`"button`" role=`"tab`" aria-controls=`"nav-$($Tab.id)`" aria-selected=`"false`">$($Tab.name)</button>"
        tabPane=@"
                <div class="tab-pane tab-pane-main" id="nav-$($Tab.id)" role="tabpanel" aria-labelledby="nav-$($Tab.id)-tab">
                <div class="card">
                    <div class="card-header">$($Tab.Title)</div>
                    <div class="card-body">$($Tab.body)</div>
                </div>
            </div>
"@
    }
    $web.Document.GetElementById("nav-tab").InnerHtml += $html.tab
    $web.Document.GetElementById("nav-tabContent").InnerHtml += $html.tabPane

    $script = $web.Document.CreateElement("script");
    $script.SetAttribute("text",$tab.scriptBlock);
    $web.Document.GetElementsByTagName("head")[0].AppendChild($script);
    $web.Document.all[$Tab.buttonID].add_click($Tab.buttonScript)
}
#===================================================================
#===================================================================
$scriptName = (Get-Item $PSCommandPath ).Basename
$configFile = "$PSScriptRoot\$scriptName.config"
$logfile = "$PSScriptRoot\$scriptName.log"
$adModulePath = "$PSScriptRoot\ActiveDirectoryModule"
$global:adModuleDllLoaded = $false
$objectIndex = 1
$global:queryIndex = 1
$global:queryResults = @{}
#===================================================================
#===================================================================
#$PSModuleAutoloadingPreference = “none” #disable auto loading of modules
loadAssembly "System.Web"
loadAssembly "System.Windows.Forms"
loadModule "ActiveDirectory" "`nTo install the Active Directory module type 'install-Module ActiveDirectory'`nYou may need to install the Remote Server Administration tools (RSAT) from Software Center`n" @{"$PSScriptRoot\Microsoft.ActiveDirectory.Management.dll"="You are running this script on a computer that does NOT have the RSAT tools install, limited functionality is available."}
$global:canExportToExcel = loadModule "ImportExcel" -required $false -commentError "`nTo enable export to Excel please install the Import Excel module type 'install-Module ImportExcel'"
#loadModule "Microsoft.PowerShell.Utility" "`nTo install the Microsoft.PowerShell.Utility module type 'install-Module Microsoft.PowerShell.Utility'"
#Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
#===================================================================
$maxResults = 1000
$global:ADObjects = @{objects=@{}}
$global:config = New-Object -TypeName PSCustomObject
$global:domains = New-Object -TypeName PSCustomObject
if(Test-Path -Path $configFile -PathType Leaf){$global:config = Get-Content($configFile) | ConvertFrom-Json}
@("properties";"groupPolicies") | %{if(!$global:config.$_){Add-Member -InputObject $global:config -type NoteProperty -Name $_ -Value (New-Object -TypeName PSCustomObject) -Force}}
#===================================================================
$html = Get-Content "$PSScriptRoot\GUI.html" -Raw
$Properties_Common =@("Name:checked";"description";"DistinguishedName:checked";"whencreated";"whenchanged")
$objects = [ordered]@{
    user = @{
        properties=$Properties_Common + @("SID:checked";"SamAccountName";"UserPrincipalName:checked";"GivenName:checked";"Surname:checked";"lastLogon:checked";"lastLogonTimestamp";"LockedOut";"Enabled:checked";"EmployeeID";"EmailAddress:checked";"memberOf:checked";"HomeDrive:checked";"HomeDirectory:checked";"PasswordExpired";"PasswordNeverExpires";"PasswordLastSet")
    }
    computer =@{
        properties=$Properties_Common + @("SID:checked";"SamAccountName";"isCriticalSystemObject:checked";"IPv4Address:checked";"DNSHostName:checked";"OperatingSystem:checked";"OperatingSystemHotfix";"OperatingSystemServicePack";"OperatingSystemVersion";"lastLogon:checked";"lastLogonTimestamp";"LockedOut";"Enabled:checked";"memberOf:checked")
    }
    group =@{
        properties=$Properties_Common + @("SID:checked";"GroupCategory";"GroupScope";"ManagedBy:checked";"Members:checked")
    }
    organizationalunit =@{
        properties=$Properties_Common + @("gPLink:checked";"isCriticalSystemObject";"ManagedBy:checked")
        options=@{ManagedBy=@{text="Enumerate Manager Names";tooltip="Looks up the user from the distinguishedName. Depending on the search filter there may be many results and this may take a long time.  Alternately you can enumerate individual ManagedBy attributes when the search has completed."}}
    }
    object=@{properties=$Properties_Common}
    trust=@{properties=@("Name:checked";"Direction:checked";"Source:checked";"Target:checked";"TrustType";"UplevelOnly";"UsesAESKeys";"UsesRC4Encryption";"DisallowTransivity";"DistinguishedName";"ForestTransitive";"IntraForest";"IsTreeParent";"IsTreeRoot";"SelectiveAuthentication";"SIDFilteringForestAware";"SIDFilteringQuarantined")}
    forest=@{properties=@("Name:checked";"Domains:checked";"DomainNamingMaster:checked";"ForestMode:checked";"GlobalCatalogs";"RootDomain";"SchemaMaster";"Sites";"SPNSuffixes";"UPNSuffixes")}
    replicationsite=@{properties=$Properties_Common + @("InterSiteTopologyGenerator")}
    replicationsubnet=@{properties=$Properties_Common + @("Location";"Site")}
}
#============================================
$OBJECTPROPERTIES = ""
$OBJECTOPTIONS = ""
$TYPELIST = ""
$objects.keys | % {
    $display=if($_ -eq 'user'){""}else{'d-none'};
    $selected=if($_ -eq 'user'){"selected"}else{''};
    $type = $_
    $TYPELIST += "<option $selected>$((Get-Culture).TextInfo.ToTitleCase($_))</option>"

    $OBJECTPROPERTIES += "<div id=`"properties-$_`" ourType=`"$_`" class=`"row div-properties $display`">"
    $objects.$_.properties | %{$props=$_.split(':'); $OBJECTPROPERTIES += "<div class=`"col-6 col-sm-3 p-2`"><div class=`"form-check`"><input type=`"checkbox`" property=`"$($props[0])`" $($props[1]) class=`"form-check-input property-$type`" id=`"$type-check-$($props[0])`"><label class=`"form-check-label`" for=`"$type-check-$($props[0])`">$($props[0])</label></div></div>" } 
    $OBJECTPROPERTIES += "</div>"

    if($objects.$_.options){
        $OBJECTOPTIONS += "<div id=`"options-$_`" ourType=`"$_`" class=`"row div-options $display`">"
        $objects.$_.options.keys | %{
            $opt = $objects.$type.options.$_
            $OBJECTOPTIONS += "<div class=`"col-6 col-sm-3 p-2`"><div class=`"form-check`"><input type=`"checkbox`" option=`"$_`" $($opt.checked) class=`"form-check-input option-$type`" id=`"$type-option-check-$_`"><label data-toggle=`"tooltip`" data-placement=`"right`" title=`"$($opt.tooltip)`" class=`"form-check-label`" for=`"$type-option-check-$_`">$($opt.text)</label></div></div>" 
        } 
        $OBJECTOPTIONS += "</div>"
    }else{
        $OBJECTOPTIONS += "<div id=`"options-$_`" ourType=`"$_`" class=`"row div-options $display`"><div class=`"alert alert-info`" role=`"alert`">There are no additional search options for $((Get-Culture).TextInfo.ToTitleCase($_)) </div></div>"
    }
}
#============================================
$global:aboutInfo = if($global:adModuleDllLoaded){'<div class="alert alert-warning" role="alert">This script is running using the Active Directory dll import, limited functionality and properties are available.<p>For full functionality please install the RSAT tools</p></div>'}else{""}
if(!$global:canExportToExcel){$global:aboutInfo += "<div class=`"alert alert-info`" role=`"alert`">The ImportExcel module is not installed.  To enable export to Excel you can install this by <a href=`"#`" onclick=`"`$('#powershellButton').attr('object','importExcel').attr('cmd','install').trigger('click')`">Clicking here</a></div>"}
#============================================
$html = $html -replace "{{TYPELIST}}", $TYPELIST
$html = $html -replace "{{OBJECTPROPERTIES}}", $OBJECTPROPERTIES
$html = $html -replace "{{OBJECTOPTIONS}}", $OBJECTOPTIONS
#============================================
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.bounds.height

function adjustClassOnHTMLElement($AddOrRemove,$className,$elementID)  {
    $classes = $web.Document.GetElementById($elementID).GetAttribute("className")
    if($classes -like "*$className*" -and $AddOrRemove -eq "Remove"){
        $web.Document.GetElementById($elementID).SetAttribute("className", $classes.replace($className,""))
    }elseif($classes -Notlike "*$className*" -and $AddOrRemove -eq "Add"){
        $web.Document.GetElementById($elementID).SetAttribute("className", $classes + " $className")
    }    
}
  
function getHomeCreds(){
    $username = $web.Document.all["home-username"].InnerText
    if($username){
        [securestring]$password =  ConvertTo-SecureString $web.Document.all["home-password"].InnerText -AsPlainText -Force
        new-object -typename System.Management.Automation.PSCredential($username,$password)
    }else{$null} 
}
function homeSubmitSearch() {
    write-log "INFO" "Initializing homeSubmitSearch" $true   
    adjustClassOnHTMLElement "Remove" "d-none" "ReadyOverlay";[System.Windows.Forms.Application]::DoEvents() 
    $err = ""
    $creds = getHomeCreds
    $searchType = $web.Document.all["home-searchType"].InnerText
    $typeSelect = $web.Document.DomDocument.getElementById("home-typeSelect") | where { $_.selected } | % { $_.text }
    $identity = $web.document.all["home-identity"].GetAttribute("value")
    $selectedDomains = $web.Document.DomDocument.getElementById("home-domainSelect") | where { $_.selected } | % { $_.text }
    $properties = @();$options = @()
    write-host "property-$typeSelect"
    $web.Document.DomDocument.getElementById("property-check").getElementsByClassName("property-$($typeSelect.tolower())") | %{
        if($_.checked){$properties += $($_.GetAttribute('property'))}
    }
    $web.Document.DomDocument.getElementById("option-check").getElementsByClassName("option-$($typeSelect.tolower())") | %{
        if($_.checked){$options += $($_.GetAttribute('option'))}
    }
    if($properties.Count -eq 0) { $err += "You must select at least 1 $typeSelect property`n" }
    if (!$identity) { $err += "You must enter an identity`n" }
    if (!$selectedDomains.count) { $err += "You must select at least 1 domain`n" }

    if ($err) {
        adjustClassOnHTMLElement "Add" "d-none" "ReadyOverlay"
        $web.Document.InvokeScript("showAlert", @("error"; "More Information Required"; $err));
    }else {
        @('typeSelect';'identity';'searchType') | %{Add-Member -InputObject $global:config -type NoteProperty -Name $_ -Value  (Get-Variable "$_" -ValueOnly -ea silentlycontinue) -Force}
        Add-Member -InputObject $global:config.properties -type NoteProperty -Name $typeSelect -Value $properties -Force
        $ReturnObjects = @{}
        #============================================
        function scriptbox($typeSelect,$searchType,$identity,$domain,$properties,$options,$ReturnObjects,$creds,$adModulePath,$maxResults){
            write-log "INFO" "starting Job for $domain"

            $params = @{Server=$domain;ResultSetSize=$maxResults;properties=$properties}  
            switch($searchType){
                "Identity"{$params.Filter = "Name -like '$identity'"}
                "Filter"{$params.Filter = $identity}
                "LDAPFilter"{$params.LDAPFilter = $identity}
            }
            if($creds){$params.credentials=$creds}
            try{
                switch ($typeSelect.toLower()) {
                    {$("trust";"replicationsite";"replicationsubnet").contains($_)} {$params.remove("ResultSetSize");break}
                    "forest" {@('ResultSetSize';'Filter';'properties') | %{$params.remove($_)};break}
                }
                $paramString = hashToString $params $true
                if($searchType -eq "RAW"){
                    $identity = $identity -ireplace "{{domain}}", $domain
                    $global:config.domains.PSObject.Properties | %{$identity = $identity -ireplace "{{$($_.name)-domains}}",  "'$($_.value -join "';'")'"}

                    write-log "INFO" "Calling Raw Expression: $identity"
                    [System.Windows.Forms.Application]::DoEvents()
                    $object = Invoke-Expression $identity
                    if($object){
                        $propObj = if($object.gettype().name -eq 'Object[]'){$object[0]}else{$object}
                        $properties = $propObj.PSObject.Properties | Select-Object -Property Name | %{$_.name}
                    }
                }else{
                    write-log "INFO" "Calling: get-ad$typeSelect $paramString"
                    [System.Windows.Forms.Application]::DoEvents()
                    $object = & "Get-AD$typeSelect" @params | select $properties
                }
                <#switch ($typeSelect) {
                    "User" {$object = Get-ADUser @params | select $properties}
                    "Computer" {$object = Get-ADComputer @params | select $properties}
                    "Group" {$object = Get-ADGroup @params | select $properties}
                    "OrganizationalUnit" {$object = Get-ADOrganizationalUnit @params | select $properties}
                    Default {Get-ADObject @params | select $properties}
                }#>

                $count = if(!$object){0}elseif($object.count){$object.count}else{1}
                write-log "INFO" "$domain Query returned $count objects"

                $result = @()
                $translations = @{memberof="Groups"}
                $valuesAreDNs = @('memberof')
                $convertProps = "lastLogon|lastLogonTimestamp|memberOf|members|gPLink|ManagedBy|Domains|GlobalCatalogs|Sites|UPNSuffixes|SPNSuffixes"
                if($count -gt 0 -and  $properties -match "\b($convertProps)\b"){ 
                     write-log "INFO" "Properties contain values that must be converted"
                    $object = $object | %{
                        $item = $_
                        $exportItem = $_.PsObject.Copy()

                        $_.PSObject.Properties | %{
                            #if($properties.contains($_.Name)){
                                #write-host "$($_.name) = $($_.TypeNameOfValue)"
                                if($_.TypeNameOfValue -eq "Microsoft.ActiveDirectory.Management.ADPropertyValueCollection"){
                                    $name = if($translations.keys.contains($_.Name.tolower())){$translations.($_.Name.tolower())}else{$_.Name}
                                    $isDN = $valuesAreDNs.contains($_.Name.tolower())
                                    if($_.Value.count -eq 0) {
                                        Add-Member -InputObject $exportItem -type NoteProperty -Name "membersCount" -Value 0 -Force
                                        Add-Member -InputObject $item -type NoteProperty -Name "membersCount" -Value 0 -Force                          
                                        Add-Member -InputObject $exportItem -type NoteProperty -Name $_.Name -Value "Empty" -Force
                                        Add-Member -InputObject $item -type NoteProperty -Name $_.Name -Value "Empty" -Force
                                    }else{
                                        switch($_.Name.tolower()){
                                            "members"{
                                                $memberCount = $item.$_.count #($item.$_|Get-Member -Type NoteProperty).count
                                                Add-Member -InputObject $exportItem -type NoteProperty -Name "membersCount" -Value $memberCount -Force
                                                Add-Member -InputObject $item -type NoteProperty -Name "membersCount" -Value $memberCount -Force                          
                                                $allResolved = $true
                                                $members = $item.$_ | %{
                                                    if($global:ADObjects.objects.containsKey($_)){
                                                        $object = $global:ADObjects.objects.$_
                                                        $resolved = if($object.givenName){"$($object.givenName) $($object.Surname) ($($object.cn))"}else{"$($object.cn)"}
                                                        "$_::$resolved"
                                                    }else{$_;$allResolved=$false}
                                                }          
                                                
                                                Add-Member -InputObject $exportItem -type NoteProperty -Name $_ -Value (($members | Sort-Object) -join "|") -Force  
                                                if($allResolved -eq $false){
                                                    $objectIndex++
                                                    Add-Member -InputObject $item -type NoteProperty -Name $_ -Value "<button type=`"button`" class=`"btn btn-link`" onClick=`"showObjects('$(($members | Sort-Object) -join "|")','Group Members',true)`">Members</button><button type=`"button`" class=`"btn btn-link`" id=`"GetAll-$objectIndex`" allObjects=`"$($members -join "|")`" title=`"Members`" onClick=`"`$('#powershellButton').attr('object','GetAll-$objectIndex').attr('cmd','enumAll').attr('resultIndex','$global:queryIndex').attr('itemIndex','$($result.count)').trigger('click')`">Enumerate Members</button>" -Force
                                                }else{
                                                    Add-Member -InputObject $item -type NoteProperty -Name $_ -Value "<button type=`"button`" class=`"btn btn-link`" onClick=`"showObjects('$(($members | Sort-Object) -join "|")','Group Members',true)`">Members</button>" -Force
                                                }   
                                                break;
                                            }
                                            default{
                                                #write-host $_
                                                Add-Member -InputObject $exportItem -type NoteProperty -Name $_ -Value (($item.$_ | Sort-Object) -join "|") -Force
                                                Add-Member -InputObject $item -type NoteProperty -Name $_ -Value "<button type=`"button`" class=`"btn btn-link`" onClick=`"showObjects('$(($item.$_ | Sort-Object) -join "|")','$_', $("$isDN".tolower()))`">$($item.$_.count) $name</button>" -Force
                                            }
                                        }    
                                    }
                                }
                            #}
                        }
                        if($properties.contains('lastLogon')){
                            $ll = if($_.lastLogon){[DateTime]::FromFileTime($_.lastLogon)}else{'Never'}
                            Add-Member -InputObject $exportItem -type NoteProperty -Name "lastLogon" -Value $ll -Force
                            Add-Member -InputObject $item -type NoteProperty -Name "lastLogon" -Value $ll -Force
                        }
                        if($properties.contains('lastLogonTimestamp')){
                            $ll = if($_.lastLogonTimestamp){[DateTime]::FromFileTime($_.lastLogonTimestamp)}else{'Never'}
                            Add-Member -InputObject $exportItem -type NoteProperty -Name "lastLogonTimestamp" -Value $ll -Force
                            Add-Member -InputObject $item -type NoteProperty -Name "lastLogonTimestamp" -Value $ll -Force
                        }
                        if($properties.contains('ManagedBy') -and $_.ManagedBy){
                            $managedBy = $_.ManagedBy
                            Add-Member -InputObject $exportItem -type NoteProperty -Name "ManagedBy" -Value $managedBy -Force
                            if($global:ADObjects.objects.containsKey($_.ManagedBy)){
                                $manager = $global:ADObjects.objects.($_.ManagedBy)
                                write-log "INFO" "Manager $($manager.givenName) $($manager.Surname) has already been found $ManagedBy"
                                $managedBy = if($manager.EmailAddress){"<a href=`"mailto:$($manager.EmailAddress)`">$($manager.givenName) $($manager.Surname) $(if($manager.EmployeeID){"($($manager.EmployeeID))"})</a>"}               
                                else{"$($manager.givenName) $($manager.Surname)"}
                                Add-Member -InputObject $exportItem -type NoteProperty -Name "ManagedBy" -Value  "$($manager.givenName) $($manager.Surname)" -Force
                            }else{
                                if($options.contains('ManagedBy')){
                                    try{
                                        $domain = $_.ManagedBy.Substring($_.ManagedBy.IndexOf("DC=")+3).Replace(",DC=",".")
                                        $paramsManagedBy = @{Server=$domain;identity=$_.ManagedBy;properties=@("givenName";"Surname";"EmployeeID";"EmailAddress")}  
                                        if($creds){$paramsManagedBy.credentials=$creds}
                                        write-log "INFO" "$domain Calling: get-adUser $(hashToString $paramsManagedBy $true)"
                                        [System.Windows.Forms.Application]::DoEvents()
                                        $manager = Get-ADUser @paramsManagedBy
                                        Add-Member -InputObject $exportItem -type NoteProperty -Name "ManagedBy" -Value  "$($manager.givenName) $($manager.Surname)" -Force
                                        $managedBy = if($manager.EmailAddress){"<a href=`"mailto:$($manager.EmailAddress)`">$($manager.givenName) $($manager.Surname) $(if($manager.EmployeeID){"($($manager.EmployeeID))"})</a>"}               
                                                    else{"$($manager.givenName) $($manager.Surname)"}
                                        $global:ADObjects.objects.($_.ManagedBy) = $manager
                                    }catch{ write-log "ERROR"  "Failed to get ADUser for $($_.ManagedBy). $($_.Exception.Message)"}
                                }else{
                                    $objectIndex++
                                    $ManagedBy = "<button type=`"button`" class=`"btn btn-link`" id=`"GetUser-$objectIndex`" userDN=`"$($_.ManagedBy)`" onClick=`"`$('#powershellButton').attr('object','GetUser-$objectIndex').attr('cmd','enumUser').attr('resultIndex','$global:queryIndex').attr('itemIndex','$($result.count)').trigger('click')`">$($_.ManagedBy)</button>"
                                }              
                             }
                             Add-Member -InputObject $item -type NoteProperty -Name "ManagedBy" -Value $managedBy -Force
                        }
                        if($properties.contains('gPLink')){
                            if($_.gPLink -like "" -or $_.gPLink -like " "){
                                Add-Member -InputObject $exportItem -type NoteProperty -Name "gpLinkCount" -Value 0 -Force
                                Add-Member -InputObject $exportItem -type NoteProperty -Name "gPLink" -Value "No Links" -Force
                                Add-Member -InputObject $item -type NoteProperty -Name "gpLinkCount" -Value 0 -Force
                                Add-Member -InputObject $item -type NoteProperty -Name "gPLink" -Value "No Links" -Force
                            }else{
                                $gpnames =@()
                                $gplinks = $_.gPLink.trimStart("[").trimEnd("]").split("][")
                                Add-Member -InputObject $exportItem -type NoteProperty -Name "gpLinkCount" -Value $gplinks.count -Force
                                Add-Member -InputObject $item -type NoteProperty -Name "gpLinkCount" -Value $gplinks.count -Force
                                write-log "INFO" "Enumerating gplink objects ($($gplinks.count))"
                                $gplinks | %{
                                    $linkArray = $_.split(';')
                                    $link = $linkArray[0].trimStart("LDAP://")
                                    $status = $linkArray[1].replace("0","Enabled,Not Enforced").replace("1","Not Enabled,Not Enforced").replace("2","Enabled,Enforced").replace("3","Enforced,Not Enabled")
                                    if($global:config.groupPolicies.PSobject.Properties.Name -contains $link){
                                        #write-log "INFO" "$global:config.groupPolicies.$link"
                                        $gpnames +=  "$link::$($global:config.groupPolicies.$link) ($status)"
                                    }else{
                                        $domain = $link.Substring($link.IndexOf("DC=")+3).Replace(",DC=",".")
                                        $paramsgpLinks = @{Server=$domain;identity=$link;properties=@("displayName")}  
                                        if($creds){$paramsgpLinks.credentials=$creds}
                        
                                        $paramString = hashToString $paramsgpLinks $true
                                        write-log "INFO" "Calling: get-adobject $paramString"
                                        try{
                                            [System.Windows.Forms.Application]::DoEvents() 
                                            $displayname = (get-adobject @paramsgpLinks | select "displayname").displayname
                                            if(!$displayname){$displayname = $link}
                                            $gpnames += "$link::$displayname ($status)"
                                            Add-Member -InputObject $global:config.groupPolicies -type NoteProperty -Name $link -Value $displayname -Force
                                        }
                                        catch{
                                            write-log "ERROR"  "Failed to get displayName for $link. $($_.Exception.Message)"
                                            $gpnames += "$link::$link ($status)"
                                            Add-Member -InputObject $global:config.groupPolicies -type NoteProperty -Name $link -Value $link -Force
                                         
                                        }
                                    }
                                }                                 
                                Add-Member -InputObject $exportItem -type NoteProperty -Name "gPLink" -Value (($gpnames | Sort-Object) -join "|") -Force
                                Add-Member -InputObject $item -type NoteProperty -Name "gPLink" -Value "<button type=`"button`" class=`"btn btn-link`" onClick=`"showObjects('$(($gpnames | Sort-Object) -join "|")','Linked Group Policies',false)`">Group Policies</button>" -Force
                            }
                        }
                        $item   
                        $result+= $exportItem
                    }
                }
                if($result){$global:queryResults."$global:queryIndex" = @{result=$result;query="$typeSelect objects with $($searchType):$identity in $domain"}}
                $ReturnObjects.$domain = $object    
            }catch{
                write-log "ERROR" "$domain. $($_.Exception.Message)"
                $ReturnObjects.$domain = $_.Exception.Message  
            } 
        }  
        #============================================     
        $heightSet = $false
        $resultCount = 0
        $errorString = ""
        $selectedDomains | %{
            $domain  = $_
            scriptbox $typeSelect $searchType $identity $_ $properties $options $ReturnObjects $creds $adModulePath $maxResults 
            $result = $ReturnObjects.$domain
            
            write-host $result.GetType().name
            if($result -eq $null){
                write-log "WARN" "$domain query returned 0 objects"
                $web.Document.InvokeScript("addTab", @("home-result";$typeSelect;$domain;"0 objects returned",'info',$global:queryIndex))
            }elseif($result.GetType().name -eq "String"){
                $web.Document.InvokeScript("addTab", @("home-result";$typeSelect;$domain;[string]$result,'danger',$global:queryIndex))
                $errorString = [string]$result
            }else{           
                $count = if([string]$result.gettype() -eq "System.Object[]"){$result.count}else{1}  
                $resultCount += $count
                write-log "INFO" "$domain returned $count objects" 
                
                $convertParams = @{Fragment = $true} #;PreContent = "<section><h2>$($domain)</h2>";#PostContent = "</section>"}             
                write-log "INFO" "Converting the results to a table" $true
                $table = $result | ConvertTo-Html @convertParams
                $table = $table -replace '<table>',"<table id=`"$domain`" class=`"display report-table`">"
                $table = $table -replace '<tr><th>',"<thead><tr><th>"
                $table = $table -replace '</th></tr>',"</th></tr></thead>"
                $table = [System.Web.HttpUtility]::HtmlDecode($table)
                write-log "INFO" "Adding the table to the GUI" $true
                $web.Document.InvokeScript("addTab", @("home-result";$typeSelect;$domain;[string]$table,'table',$global:queryIndex,($result.count -eq $maxResults)))
            }
        }
        if($resultCount -eq 0){$web.Document.InvokeScript("showAlert", @("warning"; "No Data"; if($errorString){$errorString}else{"No $typeSelect objects were found with the search string $identity"}));}
        adjustClassOnHTMLElement "Remove" "d-none" "home-result-card"
        adjustClassOnHTMLElement "Add" "show" "home-result"
        $web.Document.GetElementById("home-result-card").Focus()
        adjustClassOnHTMLElement "Add" "d-none" "ReadyOverlay"
        $global:queryIndex++
        write-host "Search complete"                  
    }   
}
function enumAll($object){
    $psBtn = $web.Document.GetElementById('powershellButton')
    $element =  $web.Document.GetElementById($object)
    $result = $global:queryResults."$($psBtn.GetAttribute("resultIndex"))".result
    $itemIndex = [int]$psBtn.GetAttribute("itemIndex")

    $parent = $element.parent
    $element.parent.innerhtml = 'Please Wait...'
    $creds = getHomeCreds
    $allObjects = $element.GetAttribute("allObjects").split("|")
    $enum  = @()
    $allObjects | %{
        $DN = $_
        if($global:ADObjects.objects.containsKey($_)){
            $object = $global:ADObjects.objects.$_
            $resolved = if($object.givenName){"$($object.givenName) $($object.Surname) ($($object.cn))"}else{"$($object.cn)"}
            $enum += "$_::$resolved"
        }else{
            $domain = $_.Substring($_.IndexOf("DC=")+3).Replace(",DC=",".")                             
            $params = @{Server=$domain;identity=$_;properties=@("ObjectClass";"cn")}  
            if($creds){$params.credentials=$creds}
            write-log "INFO" "$domain Calling: get-ADObject $(hashToString $params $true)" $true
            try{
                [System.Windows.Forms.Application]::DoEvents()
                $object = get-ADObject @params
                $resolved = $object.cn
                $global:ADObjects.objects.$DN = $object 
                switch($object.ObjectClass){
                    "user" {
                        $params.properties=@("givenName";"Surname";"cn")
                        write-log "INFO" "$domain Calling: get-ADUser $(hashToString $params $true)" $true
                        [System.Windows.Forms.Application]::DoEvents()
                        $object = get-ADUser @params
                        $resolved = if($object.givenName){"$($object.givenName) $($object.Surname) ($($object.cn))"}else{"$($object.cn)"} 
                        write-host $resolved
                        $global:ADObjects.objects.$DN = $object 
                    }
                }  
                $enum += "$DN::$resolved"
            }
            catch{
                write-log "ERROR" "Couldnt enumerate $DN. $($_.Exception.Message)"
                $enum += "$DN"
            }
        }
    }    
    Add-Member -InputObject $result[$itemIndex] -type NoteProperty -Name $element.GetAttribute("title") -Value ($enum -join "|") -Force
    $button = "<button type=`"button`" class=`"btn btn-link`" onClick=`"showObjects('$($enum -join "|")','$($element.GetAttribute("title"))',false)`">$($element.GetAttribute("title"))</button>"
    $parent.innerhtml = $button
}
function enumUser($object){
    $psBtn = $web.Document.GetElementById('powershellButton')
    $element =  $web.Document.GetElementById($object)
    $result = $global:queryResults."$($psBtn.GetAttribute("resultIndex"))".result
    $itemIndex = [int]$psBtn.GetAttribute("itemIndex")

    $innerhtml = $element.innerhtml
    $creds = getHomeCreds
    $userDN = $element.GetAttribute("userDN")
    if($global:ADObjects.objects.containsKey($_)){
        $user = $global:ADObjects.objects.$_
        $userhtml = if($user.EmailAddress){"<a href=`"mailto:$($user.EmailAddress)`">$($user.givenName) $($user.Surname) $(if($user.EmployeeID){"($($user.EmployeeID))"})</a>"}               
            else{"$($user.givenName) $($user.Surname)"}
        Add-Member -InputObject $result[$itemIndex] -type NoteProperty -Name "ManagedBy" -Value "$($user.givenName) $($user.Surname)" -Force
        $element.parent.innerhtml = $userhtml   
    }else{
        $domain = $userDN.Substring($userDN.IndexOf("DC=")+3).Replace(",DC=",".")                             
        $params = @{Server=$domain;identity=$userDN;properties=@("givenName";"Surname";"EmployeeID";"EmailAddress")}  
        if($creds){$params.credentials=$creds}
        write-log "INFO" "$domain Calling: get-adUser $(hashToString $params $true)" $true
        try{
            $element.innerhtml = 'Please Wait...'
            [System.Windows.Forms.Application]::DoEvents()
            $user = Get-ADUser @params
            $userhtml = if($user.EmailAddress){"<a href=`"mailto:$($user.EmailAddress)`">$($user.givenName) $($user.Surname) $(if($user.EmployeeID){"($($user.EmployeeID))"})</a>"}               
            else{"$($user.givenName) $($user.Surname)"}
            Add-Member -InputObject $result[$itemIndex] -type NoteProperty -Name "ManagedBy" -Value "$($user.givenName) $($user.Surname)" -Force
            $element.parent.innerhtml = $userhtml   
            $global:ADObjects.objects.$userDN = $user  
        }
        catch{
            $element.innerhtml =$innerhtml
            $web.Document.InvokeScript("showAlert", @("error"; "Failed to Enumerate User";  $_.Exception.Message))
        }
    }
}
function export($object){
    write-log "INFO" "Initializing export"
    $psBtn = $web.Document.GetElementById('powershellButton')
    $file = $psBtn.GetAttribute("file")
    $ignoreExists = $psBtn.GetAttribute("ignoreExists")
    $ext = Split-Path -Path $file -Extension
    if($ext.tolower() -ne ".xlsx" -and $global:canExportToExcel){$file = "$($file.trimEnd($ext)).xlsx"}
    if($ext.tolower() -ne ".csv" -and !$global:canExportToExcel){$file = "$($file.trimEnd($ext)).csv"}

    if($ignoreExists -ne $true -and (Test-Path -Path $file -PathType Leaf)){
        $web.Document.InvokeScript("askQuestion", @("warning";"Overwrite File?", "The file '$file' exists, do you want to overwrite it?";"`$('#powershellButton').attr('ignoreExists',true).trigger('click');"))
        return
    }  
    $psBtn.SetAttribute("ignoreExists",$false) 
    adjustClassOnHTMLElement "Remove" "d-none" "ReadyOverlay";[System.Windows.Forms.Application]::DoEvents() 
    $result = [pscustomobject]($global:queryResults."$object".result)
    if($global:canExportToExcel){
        $exported = exportToExcel $result $file "Powershell GUI" -format (@{title="Powershell GUI Export";subtitle="$($result.count) objects returned from query '$($global:queryResults."$object".query)'. Generated $(Get-Date -Format "dddd dd/MM/yyyy HH:mm")"})
        if([string]$exported.getType() -eq "string"){
            $web.Document.InvokeScript("showAlert", @("error"; "Failed to export data";$exported));
        }else{
            $web.Document.InvokeScript("askQuestion", @("success"; "Export Complete";"Successfully exported to '$file', do you want to open it?";"`$('#powershellButton').attr('object','$($file.replace("\","|"))').attr('cmd','openFile').trigger('click');"));
        }
    }else{
        try{
            [pscustomobject]$result | Export-Csv -Path $file -NoTypeInformation
            $web.Document.InvokeScript("askQuestion", @("success"; "Export Complete";"Successfully exported to '$file', do you want to open it?";"`$('#powershellButton').attr('object','$($file.replace("\","|"))').attr('cmd','openFile').trigger('click');"));
        }catch{$web.Document.InvokeScript("showAlert", @("error"; "Failed to export data"; $_.Exception.Message))}
    }
    adjustClassOnHTMLElement "Add" "d-none" "ReadyOverlay";[System.Windows.Forms.Application]::DoEvents() 
}
function openFile($object){
    write-log "INFO" "Initializing openFile"
    $file = $object.replace("|","\")
    if(!(Test-Path -Path $file -PathType Leaf)){
        write-log "ERROR" "The file '$file' doesnt exist."
        $web.Document.InvokeScript("showAlert", @("error"; "Failed to open File";"The file '$file' could not be found"))
        return
    }
    Invoke-item $file
}
function installModule($module){
    $psBtn = $web.Document.GetElementById('powershellButton')
    $scope = $psBtn.GetAttribute("scope")
    write-host $scope
    if(!$scope){ 
        $web.Document.InvokeScript("askQuestion", @("question"; "Module Scope";"Do you want to install the module '$module' for all users? Press no to install for you only.";"`$('#powershellButton').attr('object','$module').attr('scope','AllUsers').attr('cmd','install').trigger('click');";"`$('#powershellButton').attr('object','$module').attr('scope','CurrentUser').attr('cmd','install').trigger('click');"));
        return
    }
    try{
        adjustClassOnHTMLElement "Remove" "d-none" "ReadyOverlay";[System.Windows.Forms.Application]::DoEvents() 
        Install-Module -Name $module -force -scope $scope
        $web.Document.InvokeScript("showAlert", @("success"; "The module '$module' was successfully installed"))
        if($module.tolower() -eq "importexcel"){$global:canExportToExcel = $true}
    }catch{$web.Document.InvokeScript("showAlert", @("error"; "Failed to install '$module'"; $_.Exception.Message))}
    $web.Document.DomDocument.GetElementById('powershellButton').removeAttribute("scope")
    adjustClassOnHTMLElement "Add" "d-none" "ReadyOverlay";[System.Windows.Forms.Application]::DoEvents() 
}
function addDomain($domainName){
    $el = $web.Document.GetElementById("home-domainSelect")
    $option = "<option class=`"user-specified`" selected>$domainName</option>"
    if ($global:domains.PSobject.Properties.Name -contains "User Specified"){
        $global:domains."User Specified" += $domainName
        $el = $web.Document.GetElementById("User Specified")
        write-host $el.gettype()
    }else{
        Add-Member -InputObject $global:domains -type NoteProperty -Name "User Specified" -Value @($domainName) -Force
        $option  = "<optgroup id=`"User Specified`" label=`"User Specified`">$option </optgroup>"
    }
    $el.InnerHtml = $el.InnerHtml + $option
    adjustClassOnHTMLElement "Remove" "d-none" "btn-remove-domain";[System.Windows.Forms.Application]::DoEvents() 
}
function removeDomain($data){
    $vals = $data -split ","
    if($global:domains.($vals[1])){
        write-host "removing $($vals[0]) from $($vals[1])"
        $domains = @()
        $global:domains.($vals[1]) | %{ if($_ –ne $vals[0]){$domains+=$_} }
        if($domains.count){Add-Member -InputObject $global:domains -type NoteProperty -Name $vals[1] -Value $domains -Force}
        else{$global:domains.PSObject.Properties.remove($vals[1])}
        
    }    
}
function saveQuery($name){
    $cmd = $web.document.all["home-identity"].GetAttribute("value")
    if ($global:config.PSobject.Properties.Name -contains "Saved Queries"){
        Add-Member -InputObject $global:config."Saved Queries" -type NoteProperty -Name $name -Value $cmd -Force
    }else{
        Add-Member -InputObject $global:config -type NoteProperty -Name "Saved Queries" -Value @{$name=$cmd} -Force
    }
}
function ShowSavedQueries{
    $options = $global:config."Saved Queries".PSobject.Properties | %{"<option value=`"$([System.Web.HttpUtility]::HtmlEncode($_.Value))`">$($_.Name)</option>"}
    $web.Document.InvokeScript("loadQuery", @(($options -join "")))
}
function domainsFromForest($data){
    $vals = $data+',,' -split ","
    $description = if($vals[1]){$vals[1]}else{$vals[0]}
    $el = $web.Document.GetElementById("home-domainSelect")
    try{
        write-log "INFO" "Calling: (Get-ADForest -Identity '$($vals[0])'').domains"
        [System.Windows.Forms.Application]::DoEvents() 
        $domains = (Get-ADForest -Identity $vals[0]).domains
        Add-Member -InputObject $global:domains -type NoteProperty -Name $description -Value $domains -Force
        $DOMAINLIST = "<optgroup id=`"$description`" label=`"$description`">"
        $domains | %{ $DOMAINLIST +="<option class=`"user-specified`">$_</option>" }
        $DOMAINLIST += "</optgroup>"
        $el.InnerHTML += $DOMAINLIST
        adjustClassOnHTMLElement "Remove" "d-none" "btn-remove-domain";[System.Windows.Forms.Application]::DoEvents() 
    }catch{
        $web.Document.InvokeScript("showAlert", @("error"; "Failed to retrieve Domains from $($vals[0])"; $_.Exception.Message))
    }
}
function DocumentCompleted(){
    write-host -Message "DocumentCompleted event fired"
    if ($global:isLoaded -eq $true) { return }
    $global:isLoaded = $true
    if($global:config.PSobject.Properties.Name -contains 'Form Dimensions'){
        $dimensions = $global:config.'Form Dimensions'.PSObject.Properties
        $form.SetDesktopLocation($dimensions['left'].Value, $dimensions['top'].Value);
    }
    #==============================================
    #Load plugins    
    Get-ChildItem -path "$PSScriptRoot\Plugins\*" -Include *.ps1 | %{$file=$_.name;try{. ("$PSScriptRoot\Plugins\$file")}catch{write-log "ERROR" "An error occurred loading the plugin file $file. $($_.Exception.Message)"}} 
    #==============================================
    #Get version and about
    try{
        write-log "INFO" "Attempting to get version info from $($app.repoRaw)$($app.versionFile)"
        [System.Windows.Forms.Application]::DoEvents() 
        $appInfo = "$(Invoke-WebRequest "$($app.repoRaw)$($app.versionFile)")||".split("|")
        if($appInfo[0]-match '[0-9].[0-9].[0-9]' -and $app.version -lt $appInfo[0]){
            $app.newVersion = $true
            try{
                [System.Windows.Forms.Application]::DoEvents() 
                write-log "INFO" "Attempting to download the updater from $($app.repoRaw)$($app.updater)"
                (New-Object System.Net.WebClient).DownloadFile("$($app.repoRaw)$($app.updater)", "$PSScriptRoot\updater.ps1")
            }catch{write-log "ERROR" "Failed to download updater script. $($_.Exception.Message)"} 

            $versionInfo = "A newer version of this app is available.<br> You have version $($app.version), the latest version is $($appInfo[0]).<p/><p>You can review the latest version info here <a href=`"#`" onClick=`"`$('#powershellButton').attr('object','$($appInfo[1])').attr('cmd','StartProcess').trigger('click')`">$($appInfo[1])</a><p>The new version features $($appInfo[2])"
            $global:aboutInfo += "<div class=`"alert alert-info`" role=`"alert`">$versionInfo</div>"
            write-log "INFO" "A newer version of this app is available.You have version $($app.version), the latest version is $($appInfo[0])"
            $web.Document.InvokeScript("askQuestion", @("info";"Newer Version Available";"$versionInfo.<p>Do you want to download the new version now?";"`$('#powershellButton').attr('cmd','doUpdate').trigger('click');"))
        }
    }catch{write-log "ERROR" "Failed to get the current version. $($_.Exception.Message)"}
    
    if(!(Test-Path -Path "$PSScriptRoot\about.md" -PathType Leaf) -or $app.newVersion){
        try{
            [System.Windows.Forms.Application]::DoEvents() 
            write-log "INFO" "Attempting to download readme file from $($app.repoRaw)$($app.readme)" -logonly $true
            (New-Object System.Net.WebClient).DownloadFile("$($app.repoRaw)$($app.readme)", "$PSScriptRoot\about.md")
        }catch{write-log "ERROR" "Failed to download readme file. $($_.Exception.Message)"}  
    }
    if(Test-Path -Path "$PSScriptRoot\about.md" -PathType Leaf){
        try{
            [System.Windows.Forms.Application]::DoEvents() 
            write-log "INFO" "loading readme info" -logonly $true
            $readme = Get-Content("$PSScriptRoot\about.md")
            $global:aboutInfo += ($readme.Split([string[]]"`r`n", [StringSplitOptions]::None) -join "<br>")
        }catch{write-log "ERROR" "Failed to download readme file. $($_.Exception.Message)"}  
    }
    $web.Document.GetElementById("home-about").InnerHTML =  $global:aboutInfo
    #==============================================
    
    write-log "INFO" "Registering form events"
    $web.document.all["home-submitSearch"].Add_click({homeSubmitSearch})

    Do{[System.Windows.Forms.Application]::DoEvents() ;Start-Sleep -Milliseconds 100}
    While ($web.document.all["main"].DomElement.clientHeight -eq 0)
    if(-not($global:config.PSobject.Properties.Name -contains 'Form Dimensions')){
        if($web.Height -lt $form.size.height){$web.Height = $form.size.height = $web.document.all["main"].DomElement.clientHeight +80}
    }
    
    $domainSelected = $null
    #load the configuration from the file
    $global:config.PSObject.Properties | %{
        $property = $_.name
        $value = $_.Value
        #write-host "$property = $value"
        switch($property){
            'domainSelected'{if($value){$domainSelected = $value} ;break}
            'domains'{if($value){$global:domains = $value} ;break}
            "properties"{
                $value.PSObject.Properties | %{
                    $type = $_.name.tolower()
                    $web.Document.DomDocument.getElementById("property-check").getElementsByClassName("property-$type") | %{
                        $_.checked = $false
                        #$_ | select *  | %{write-host $_} 
                    }
                    $_.value | %{
                        $el = $web.document.DomDocument.getElementById("$type-check-$_")
                        if($el){$el.checked=$true}
                        else{$web.Document.InvokeScript("addPropertyCheckbox", @($type.tolower();$_))}
                    }
                }
                break
            }
            'typeSelect'{
                $web.Document.InvokeScript("hometypeSelect", @($value.tolower()))
                $web.Document.DomDocument.getElementById("home-typeSelect") | %{if($_.text -eq $value){$_.selected = $true}else{$_.selected = $false}}
                break
            }
            'searchType'{
                $web.Document.InvokeScript("changeSearch", @($value))
                break
            }
            default{
                if($web.document.all["home-$property"]){$web.document.all["home-$property"].InnerText = $value}
            }
        }
    }
    $DOMAINLIST = ""
    if($global:domains.psobject.Properties.count){
        $global:domains.PSObject.Properties | %{$DOMAINLIST += "<optgroup id=`"$($_.Name)`" label=`"$($_.Name)`">"; $_.value | % {$selected = if($domainSelected -and $_ -eq $domainSelected){"selected"}else{""} $DOMAINLIST +="<option $selected>$_</option>" } }
        $DOMAINLIST += "</optgroup>"
        $web.Document.GetElementById("home-domainSelect").InnerHTML = $DOMAINLIST
        adjustClassOnHTMLElement "Remove" "d-none" "btn-remove-domain";[System.Windows.Forms.Application]::DoEvents() 
    }
    if($DOMAINLIST -eq ""){$web.Document.InvokeScript("addDomain", @("There are no domains or servers configured, you must enter a domain, server or forest (to enumerate domains)."))}

    $web.Document.GetElementById("about-header").InnerText = "$($app.title) Version $($app.version)"
    $web.Document.all['powershellButton'].add_click({
        $psBtn = $web.Document.GetElementById('powershellButton')
        $object = $psBtn.GetAttribute("object")
        switch($psBtn.GetAttribute("cmd")){
            "enumUser"{enumUser $object;break;}
            "enumAll"{enumAll $object;break;}
            "remove"{$global:queryResults.Remove("$object");break;}
            "export"{export $object;break;}
            "openFile"{openFile $object;break;}
            "install"{installModule $object;break;}
            "addDomain"{addDomain $object;break;}
            "removeDomain"{removeDomain $object;break;}
            "domainsFromForest"{domainsFromForest $object;break;}
            "saveQuery"{saveQuery $object;break;}
            "ShowSavedQueries"{ShowSavedQueries;break;}
            "doUpdate"{
                write-log "INFO" "Calling: Start-Process pwsh -ArgumentList -ep Bypass -f '$PSScriptRoot\updater.ps1'"
                Start-Process "pwsh" -WorkingDirectory $PSScriptRoot -ArgumentList "-noe","-ep","Bypass","-f","`"updater.ps1`"","AD Powershell GUI","`"$PSCommandPath`"",$app.version,"$($app.repoRaw)main/","$($app.repoRaw)$($app.versionFile)","powershellGUI.ps1;GUI.html";$form.close();break;}
            "StartProcess"{
                $argList = $psBtn.GetAttribute("argList")
                $psBtn.setAttribute("argList","")
                Start-Process $object -ArgumentList $argList
                break;
            }
            default{write-host "$_ object = $object"}
        }        
    })
    $web.Document.InvokeScript("setVariable",@("canExportToExcel";$global:canExportToExcel))
    $form.activate()
    adjustClassOnHTMLElement "Add" "d-none" "LoadingOverlay"
}
#============================================
write-log "INFO" "===============================================================================" $true 
write-log "INFO" "Initializing $($MyInvocation.MyCommand)" $true 
#============================================
#============================================
# Main
$global:isLoaded = $false
$form = New-Object -TypeName System.Windows.Forms.Form -Property  @{Width = 1200; Height = 410;StartPosition=1;MaximizeBox=$false; text = $app.title}
$web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{DocumentText = $html; ScriptErrorsSuppressed = $false}
$web.Add_DocumentCompleted({DocumentCompleted})

#write-host "========================================"
#$web | select *  | %{$_ -split ";" | %{write-host $_} }
#write-host "========================================"

write-log "INFO" "Form version = $($form.ProductName) $($form.ProductVersion)"
$form.Add_ResizeEnd({write-host "form resized"; $web.Width = $form.size.Width-40; $web.Height = $form.size.Height-60 })
$form.Add_Closing({write-host "form closing"; Add-Member -InputObject $global:config -type NoteProperty -Name "Form Dimensions" -Value @{top=$form.top;left=$form.left;width=$form.Width;height=$form.height} -Force })
if($global:config.PSobject.Properties.Name -contains 'Form Dimensions'){
    $dimensions = $global:config.'Form Dimensions'.PSObject.Properties
    $dimensions | %{$form.($_.name) = $_.value}
}
if($form.width -lt 1200){$form.width=1200}
if($form.height -lt 633){$form.width=633}
if($form.left -lt 0 -or $form.left -gt 3000){$form.left=100}
if($form.top -lt 0 -or $form.top -gt 3000){$form.top=100}

$web.width = $form.size.Width - 20; if($web.Height -lt $form.size.height){$web.Height = $form.size.height - 20}
$form.Controls.Add($web)
$form.ShowDialog() | out-null
write-log "INFO" "Closing  $($MyInvocation.MyCommand)" $true 
#ANYTHING UNDER THIS LINE WILL ONLY RUN ONCE THE WINDOW HAS BEEN CLOSED
#============================================
#export the configuration to the config file
Add-Member -InputObject $global:config -type NoteProperty -Name "domains" -Value $global:domains -Force
Add-Member -InputObject $global:config -type NoteProperty -Name "domainSelected" -Value ($web.Document.DomDocument.getElementById("home-domainSelect") | where { $_.selected } | % { $_.text }) -Force
$global:config | ConvertTo-Json -Depth 100 | Out-File $configFile
$form.Dispose()
