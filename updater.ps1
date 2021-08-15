<#
.SYNOPSIS
	This is a script updating script
.NOTES
	Author		: Hayden Trail  
    email		: hayden@tailoredit.co.nz
    Company		: Tailored IT Solutions
    File Name	: updater.ps1
#>

[cmdletbinding()]Param (
    [Parameter(Mandatory=$true)][string]$callingScript,
    [Parameter(Mandatory=$true)][string]$appVersion,
    [Parameter(Mandatory=$true)][string]$repoRaw,
    [Parameter(Mandatory=$true)][string]$versionFile,
    [Parameter(Mandatory=$true)][array]$filesToDownload
    )
$global:app = @{title="Script Updater";version="0.1";isLoaded=$false}
$afterCmd = "`$('#powershellButton').attr('object','$callingScript').attr('cmd','returnToCaller').trigger('click');"
$retryCmd = "`$('#powershellButton').attr('cmd','retry').trigger('click');"
$downloadFile = {
    Param($repoRaw,$localFolder,$file,$results)
    try{
        [System.Windows.Forms.Application]::DoEvents() 
        write-host "Attempting to download $($repoRaw)$file to $localFolder\$file"
        (New-Object System.Net.WebClient).DownloadFile("$($repoRaw)$file", "$localFolder\$file")
        $results.$file = $true
    }catch{
        write-host "Failed to download $file. $($_.Exception.Message)" -ForegroundColor "red"
        $results.$file = "Failed to download $file. $($_.Exception.Message)"
    }  
}
function proceed(){
    $jobsDone = $filesToDownload.count
    $Results = @{}
    $filesToDownload | ForEach-Object -begin {$index=0} -process{
        $listItem = $web.Document.CreateElement("li");
        $listItem.SetAttribute("className","list-group-item d-flex justify-content-between align-items-center");
        $listItem.InnerHtml = "$_<span id=`"$_`" class=`"badge bg-primary badge-pill`">Downloading...</span>"
        $web.Document.GetElementById("updater-list").AppendChild($listItem);
        
        $paramJob = @{name=$_;ScriptBlock=$downloadFile;StreamingHost=$Host;InputObject=$web;ArgumentList=@($repoRaw;$PSScriptRoot;$_;$results)}
        write-host "Calling: Start-ThreadJob for $_"
        Start-ThreadJob @paramJob
    } 
    Get-Job | %{
        $file = $_.Name
        write-host "$file Job State = $($_.State)"
        $err = $false
        Do{
            try{
                write-host "$file Job State = $($_.State); Sleeping..."
                Start-Sleep -Milliseconds 500
                [System.Windows.Forms.Application]::DoEvents()
            }catch{
                $web.Document.GetElementById($file).InnerHtml = $_.Exception.Message
                $web.Document.GetElementById($file).SetAttribute("className","badge bg-danger badge-pill")
                write-host $_.Exception.Message -ForegroundColor "red"
                $err = $true
            }                
        }
        While (!@("Completed";"Failed";"Stopped";"Suspended";"Disconnected").contains($_.State))
        write-host "$file Job State = $($_.State)"
        $jobsDone--
        if(-not($err)){
            if($results.$file -eq $true){
                $web.Document.GetElementById($file).InnerHtml = $_.State
                $web.Document.GetElementById($file).SetAttribute("className","badge bg-success badge-pill")
            }else{
                $web.Document.GetElementById($file).InnerHtml = $results.$file
                $web.Document.GetElementById($file).SetAttribute("className","badge bg-danger badge-pill")
            }
        }
        remove-job $_
        if($jobsDone -eq 0){$web.Document.GetElementById("updaterButton").SetAttribute("className", "col-3 btn btn-outline-info")}
    }   
}
function getCurrentVersion(){
    #==============================================
    #Get version
    try{
        $loader.InnerText="Attempting to get version info";[System.Windows.Forms.Application]::DoEvents() 
        write-host "Attempting to get version info from $versionFile"
        $appInfo = "$(Invoke-WebRequest $versionFile)||".split("|")
        if($appInfo[0]-match '[0-9].[0-9].[0-9]' -and $appVersion -lt $appInfo[0]){
            $fileList = $filesToDownload | %{"<li class=`"list-group-item`">$_</li>"}
            $web.Document.InvokeScript("askQuestion", @("question"; "Proceed With Download?";"You are about to overwrite the following files:<p><ul class=`"list-group list-group-flush`">$fileList</ul><p><hr>Proceed?";"`$('#powershellButton').attr('cmd','proceed').trigger('click');";$null;"`$('#powershellButton').attr('cmd','returnToCaller').trigger('click');"));
        }else{
            $versionInfo = "You have version $($app.version), the latest version is $($appInfo[0]).<p/><p>An update is not required.  Press ok to re-load the master script"
            $web.Document.InvokeScript("showAlert", @("info";"No Update Required";$versionInfo;$afterCmd))
        }
    }catch{
        write-host "Failed to get the current version. $($_.Exception.Message)" -ForegroundColor "red"
        $web.Document.InvokeScript("askQuestion", @("error"; "Failed to get the current version"; "$($_.Exception.Message)<p>Do you want to retry?.";$afterCmd;$null;$retryCmd))
    }   

}
function DocumentCompleted(){
    write-host -Message "DocumentCompleted event fired"
    if ($global:app.isLoaded -eq $true) { return }
    $global:app.isLoaded = $true
    
    $web.Document.GetElementById("Ready").SetAttribute("className", "d-none")
    $web.Document.GetElementById("Updater").SetAttribute("className", "container-fluid row d-flex justify-content-md-center")
    $loader = $web.Document.all["loadingContent"]
    [System.Windows.Forms.Application]::DoEvents() 
    #==============================================
    # register button clicks
    $web.Document.all['powershellButton'].add_click({
        $psBtn = $web.Document.GetElementById('powershellButton')
        $object = $psBtn.GetAttribute("object")
        switch($psBtn.GetAttribute("cmd")){
            "returnToCaller"{if($callingScript){Start-Process "pwsh" -ArgumentList "-ep Bypass -f `"$callingScript`""};$form.close();break;}
            "retry"{getCurrentVersion;break}
            "proceed"{proceed;break}
            default{write-host "$_ object = $object"}
        }        
    })
    getCurrentVersion
    $web.Document.GetElementById("LoadingOverlay").SetAttribute("className", "d-none")
}
$html = Get-Content "$PSScriptRoot\GUI.html" -Raw
$form = New-Object -TypeName System.Windows.Forms.Form -Property  @{Width = 1000; Height = 600;StartPosition=1;MaximizeBox=$false;SizeGripStyle=2;text = $global:app.title}
$web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{DocumentText = $html; ScriptErrorsSuppressed = $false}
$web.Add_DocumentCompleted({DocumentCompleted})
$web.width = $form.size.Width - 20; if($web.Height -lt $form.size.height){$web.Height = $form.size.height - 20}
$form.Controls.Add($web)
$form.ShowDialog() | out-null
$form.Dispose()