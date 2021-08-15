[CmdletBinding()] 
param
(	
	[Parameter(Mandatory=$false)]
	[String] $DomainName= $null,

	[Parameter(Mandatory=$false)]
	[String] $UserName= $null,

	[Parameter(Mandatory=$false)]
	[String] $Password= $null,

    [Parameter(Mandatory=$false, HelpMessage="Prompt for alternate credentials")]
    [switch]
    $AlternateCredential,

    [Parameter(Mandatory=$false, HelpMessage="Export ACL from Organizational Units")]
    [switch]
    $ExportACL,

    [Parameter(Mandatory=$false, HelpMessage="Export Organizational Units from Domain")]
    [switch]
    $ExportOU,

    [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$false, HelpMessage="Export AdminSDHolder ACL")]
    [switch]
    $AdminSDHolder,

    [Parameter(Mandatory=$false, HelpMessage="Get Organizational Unit Volumetry")]
	[Switch] $OUVolumetry,

    [Parameter(Mandatory=$false)]
	[String] $SearchRoot = $null,

    [Parameter(Mandatory=$false, HelpMessage="Name of auditor")]
	[String] $AuditorName = $env:username,
	
	[Parameter(Mandatory=$false, HelpMessage="User/Password and Domain required")]
	[Switch] $UseSSL,

    [parameter(ValueFromPipelineByPropertyName=$true, Mandatory= $false, HelpMessage= "Export data to multiple formats")]
    [ValidateSet("HTML","CSV","JSON","XML","PRINT")]
    [Array]$ExportTo=@("HTML"),

    [parameter(Mandatory= $false, HelpMessage= "Optional Scripts to run from Plugins directory")]
    [Array]$Optional=@()
)

$MyParams = $PSBoundParameters

function RunReview {

    $starttimer = Get-Date
    Write-Host "Running review"

    #loop through the form controls and add the values to the $MyParams array
    foreach($control in (get-variable WPFInput* -valueOnly)){
        SetParameter @{form=$False;Parameter=($control.name.replace("Input",""));Element=$control}  
    }
    foreach($control in (get-variable WPFExportTo* -valueOnly)){
        SetParameter @{form=$False;Parameter='ExportTo';value=($control.name.Replace("ExportTo",""));Element=$control}  
    }
    foreach($control in (get-variable WPFOptional* -valueOnly)){
        SetParameter @{form=$False;Parameter='Optional';value=($control.name.Replace("Optional",""));Element=$control}  
    }
    
    #list the parameters that the script will use now
    foreach($param in $MyParams.keys){
        Write-Host "$param = $($MyParams[$param].GetType()) = $($MyParams[$param])"
    }
    
    $WPFsettings.Visibility="Hidden"
    $WPFLog.Visibility="Visible"

    LogItem $WPFLog @{text='test red'} 
    LogItem $WPFLog @{text='test red';color="red"} $true
    ToggleElement (get-variable WPF* -valueOnly) "System.Windows.Controls.Button" "IsEnabled" $false
       
    $stoptimer = Get-Date
    LogItem $WPFLog @{text=$("Total time for JOBs: {0} Minutes" -f [math]::round(($stoptimer – $starttimer).TotalMinutes , 2))} $true    
  
    ToggleElement $WPFReset_Button  "System.Windows.Controls.Button" "Visibility"  "Visible"
    ToggleElement (get-variable WPF* -valueOnly) "System.Windows.Controls.Button" "IsEnabled"  $true
    
}
Function SetParameter([Object] $Data){
    if($Data.form){
        if($MyParams.ContainsKey($Data.Parameter)){
            if($data.Element.PSTypeNames -contains "System.Windows.Controls.CheckBox"){
                if($MyParams[$Data.Parameter].Contains($Data.Value)){$Data.Element.IsChecked = $true}
            }elseif($data.Element.PSTypeNames -contains "System.Windows.Controls.PasswordBox"){
                $Data.Element.password = $MyParams[$Data.Parameter]
            }else{       
                $Data.Element.text = $MyParams[$Data.Parameter]
            }                  
        }
    }Else{
       if($data.Element.PSTypeNames -contains "System.Windows.Controls.CheckBox"){
            if($Data.Element.IsChecked){
                if($MyParams[$Data.Parameter].length -eq 0){
                    $MyParams[$Data.Parameter] = @($Data.Value)
                }elseif(!$MyParams[$Data.Parameter].Contains($Data.Value)){
                    $MyParams[$Data.Parameter] += $Data.Value
                }
            }else{
                if($MyParams[$Data.Parameter] -contains $Data.Value) {$MyParams[$Data.Parameter] = $MyParams[$Data.Parameter] | Where-Object { $_ –ne $Data.Value }}
            }
        }elseif($data.Element.PSTypeNames -contains "System.Windows.Controls.PasswordBox"){
            $MyParams[$Data.Parameter] = $Data.Element.password
        }else{
            $MyParams[$Data.Parameter] = $Data.Element.text
        }           
    }
}
#---------------------------------------------------
# Checks if any element in array B exist in Array A
#---------------------------------------------------
function BinA([Object] $a,[Object] $b){
    return ($(Compare-Object $b $a -includeequal -excludedifferent).count -gt 0)
}
#---------------------------------------------------
# Import Modules
#---------------------------------------------------	
$ScriptPath = $PWD.Path #Split-Path $MyInvocation.MyCommand.Path -Parent
. $ScriptPath\Forms\Form_ADReview.ps1


#===========================================================================
# Use this space to add code to the various form elements in your GUI
$WPFBegin_Button.Add_Click({
    $WPFlog.Items.Clear()
    RunReview
})
$WPFCancel_Button.Add_Click({$form.Close()})
$WPFReset_Button.Add_Click({
    $WPFsettings.Visibility="Visible"
    $WPFLog.Visibility="Hidden"
})
foreach($param in $MyParams.keys){
    Write-Host "$param = $($MyParams[$param].GetType()) = $($MyParams[$param])"
}
#loop through the $MyParams array (cmd line variables) and add them to the form controls
foreach($control in (get-variable WPFInput* -valueOnly)){
    SetParameter @{form=$true;Parameter=($control.name.replace("Input",""));Element=$control}  
}
foreach($control in (get-variable WPFExportTo* -valueOnly)){
    SetParameter @{form=$true;Parameter='ExportTo';value=($control.name.Replace("ExportTo",""));Element=$control}  
}
#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null
