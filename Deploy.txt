# Author: KM

# Version: 1.10

# Usage :

# .\deploy.ps1 -btsMgmtDbName <Alias_MgMtDb_Name>

# .\deploy.ps1 -help 

 

#################################################################################

param(

[string]$btsMgmtDbServerName = "localhost",

[string]$btsMgmtDbName = "BizTalkMgmtDb",

[string]$msiRolloutSourcesRootFolder = "D:\support",

[string]$msiBackoutSourcesRootFolder = "D:\support",

[switch]$exportCSV=$false,

[switch]$backoutMode=$false,

[switch]$stopOnly=$false,

[switch]$startOnly=$false,

[switch]$validateOnly=$false,

[switch]$skipValidation=$false,

[switch]$keepHostsAndApplicationsStopped=$false,

[switch]$help=$false

)

 

#---- CU versions 

Set-Variable -name BizTalk_CU_VERSION -value "5" -option Constant

Set-Variable -name HostIntegrationServer_CU_VERSION -value "3" -option Constant

Set-Variable -name SQL_SP_VERSION -value "3" -option Constant

 

 

#---- Constants

Set-Variable -name WORKSHEET_PACKAGES -value "Packages" -option Constant

Set-Variable -name WORKSHEET_Deployment_Plan -value "Deployment_Plan" -option Constant

Set-Variable -name WORKSHEET_Servers -value "Servers" -option Constant

Set-Variable -name PROJECT_VERSION -value "1.0" -option Constant

Set-Variable -name MSBUILD_PARAMS -value "/p:Configuration=Server;SkipUndeploy=false;SkipHostInstancesRestart=true;StartApplicationOnDeploy=false;StartReferencedApplicationsOnDeploy=false;SkipIISReset=true;DeployPDBsToGac=false;DeployBizTalkMgmtDB=true;EnableAllReceiveLocationsOnDeploy=false;UndeployIISArtifacts=true;AutoTerminateInstances=false" -option Constant

Set-Variable -name DEPLOY_BAT -value "Deploy_2015.bat" -option Constant

Set-Variable -name BTDF_2015_SUFFIX -value "_2015" -option Constant

Set-Variable -name REMOTE_MSI_FOLDER -value "c:\windows\temp" -option Constant

Set-Variable -name REMOTE_LOG_FOLDER -value "DeployLogs" -option Constant

Set-Variable -name LOCAL_LOG_FOLDER -value "Results" -option Constant

 

$DataStamp = get-date -Format yyyyMMddTHHmmss

$runLogFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\{0}-{1}.log" -f "deploy",$DataStamp

$runErrLogFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\{0}-{1}.log" -f "ERRORS-deploy",$DataStamp

$runWarnLogFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\{0}-{1}.log" -f "WARNINGS-deploy",$DataStamp

 

$prgFilesLoc = ${Env:ProgramFiles(x86)}

$msBuildPath = [System.IO.Path]::Combine($env:windir,"Microsoft.NET","Framework","v4.0.30319","MSBuild.exe")

$msiSourcesRootFolder = $msiRolloutSourcesRootFolder

if($backoutMode) { $msiSourcesRootFolder = $msiBackoutSourcesRootFolder }

  

 

$configFileName="deploy.xlsx"

$installedMap = @{}

$global:packagesMap = $null

$global:deploymentPlanMap  = $null

$global:serversMap = $null

$btsHostsMap = $null

$windowsPath = $Pwd

$serversList = @{}

$settingsMap = @{}

 

function Usage() {

    Write-Host "Usage =>"

    Write-Host "Parameters: "

    Write-Host "-btsMgmtDbServerName : SQL Server Instance name or ALIAS name which hosts BizTalkMgmtDb."

    Write-Host "        localhost by default."

    Write-Host "-btsMgmtDbName : Database name of BizTalkMgmtDb."

    Write-Host "        BizTalkMgmtDb by default."

   Write-Host "-msiRolloutSourcesRootFolder : The root folder on the local server executing this script that contains ROLLOUT MSIs to be deployed. Child subfolders of this root folder MUST be named according to the BizTalk application MSI they internally contain. Ex- If -msiSourcesRootFolder passed in is 'D:\Support\Release411' the actual MSI source must be placed in D:\Support\Release411\RSA.ESB.Services.Customer\RSA.ESB.Services.Customer-1.0.0.msi"

    Write-Host "        D:\Support by default."

    Write-Host "-msiBackoutSourcesRootFolder : The root folder on the local server executing this script that contains BACKOUT MSIs to be deployed to backout a release. Child subfolders of this root folder MUST be named according to the BizTalk application MSI they internally contain. Ex- If -msiSourcesRootFolder passed in is 'D:\Support\Release411' the actual MSI source must be placed in D:\Support\Release411\RSA.ESB.Services.Customer\RSA.ESB.Services.Customer-1.0.0.msi"

    Write-Host "        D:\Support by default."

    Write-Host "-backoutMode : Switch to backout of a release and restore backout versions of MSIs and BizTalk applications."

    Write-Host "        False by default. To backout of a release and restore earlier verions of MSIs specify -backoutMode in the command line"

    Write-Host "-help : Switch to display usage reference"

    Write-Host "        False by default. To see help specify -help in the command line"

    Write-Host "-exportCSV : Switch to convert the Excel file to CSVs per work sheet. This requires the deploy.xlsx file to be in the same path as the script."

    Write-Host "        False by default. To export CSVs from Excel specify -exportCSV in the command line"

    Write-Host "-stopOnly : Only STOP all BizTalk applications and hosts and return without doing anything further."

    Write-Host "        False by default. To STOP BTS apps and hosts specify -stopOnly in the command line"

    Write-Host "-startOnly : Only START all BizTalk applications and hosts and return without doing anything further."

    Write-Host "        False by default. To START BTS apps and hosts specify -startOnly in the command line"

    Write-Host "-validateOnly : Only VALIDATE CSVs generated from deploy.xlsx as well as the release binaries/XML files available in the physical path pointed to by -msiSourcesRootFolder parameter."

    Write-Host "        False by default. To VALIDATE the release pack and return validation status without performing any deployment specify -validateOnly in the command line"

    Write-Host "-skipValidation : Deploy without performing validation assuming that the release pack has already been validated previously."

    Write-Host "        False by default. To SKIP VALIDATION during deployment specify -skipValidation in the command line"

    Write-Host "-keepHostsAndApplicationsStopped : Do NOT start BizTalk applications and hosts at the end of deployment."

    Write-Host "        True by default. To avoid starting BTS applications and hosts after deployment specify -keepHostsAndApplicationsStopped:`$false in the command line"

            Write-Host "The following command converts each worksheet in 'deploy.xlsx' to CSV files. By convention deploy.ps1 uses deploy.xlsx as the configuration file. CSV files must be exported from deploy.xlsx to be able to run in production without requiring Office installation."

    Write-Host "Example - powershell .\deploy.ps1 -exportCSV"     

    Write-Host "The following command expects CSV files of pattern *Rollout.csv & *Deployment_Plan.csv to be present in the same path as the script."     

    Write-Host "Example - powershell .\deploy.ps1 -btsMgmtDbName Alias_MgMtDb_Name -msiSourcesRootFolder d:\Rel100"             

}

 

 

 

 

function FindSheet([Object]$workbook, [string]$name)

{

    $sheetNumber = 0

    for ($i=1; $i -le $workbook.Sheets.Count; $i++) {

        if ($name -eq $workbook.Sheets.Item($i).Name) { $sheetNumber = $i; break }

    }

    return $sheetNumber

}

 

 

function SetActiveSheet([Object]$workbook, [string]$name)

{

    if (!$name) { return }

    $sheetNumber = FindSheet $workbook $name

    if ($sheetNumber -gt 0) { $workbook.Worksheets.Item($sheetNumber).Activate() }

    return ($sheetNumber -gt 0)

}

 

function Excel2CSV([string]$filePath, [string]$SheetName = "")

{

    $csvFile = Join-Path $Pwd ("{0}_{1}.csv" -f (Get-Item -path $filePath).BaseName, $SheetName)

    if (Test-Path -path $csvFile) { Remove-Item -path $csvFile }

 

    $xlCSVType = 6

    $excelObject = New-Object -ComObject Excel.Application  

    $excelObject.Visible = $false 

    $workbookObject = $excelObject.Workbooks.Open($filePath)

    SetActiveSheet $workbookObject $SheetName | Out-Null

    $workbookObject.SaveAs($csvFile,$xlCSVType) 

    $workbookObject.Saved = $true

    $workbookObject.Close()

 

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) |

        Out-Null

    $excelObject.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) |

        Out-Null

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($csvFile) |

        Out-Null

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlCSVType) |

        Out-Null

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($SheetName) |

        Out-Null

    [System.GC]::Collect()

    [System.GC]::WaitForPendingFinalizers()        

}

 

function ParsePackagesCSV($csv) {

    $csv | % { 

        $type = $_.Type

        $package = $_.Package        

        $rolloutVersion = $_.RolloutVersion

        $backoutVersion = $_.BackoutVersion

        $msi = "$package-$rolloutVersion.msi"

        $release = $_.Release

        $scope = $_.Scope        

    }    

}

 

function MapInstalledProducts() {

    LogInfo "Inferring product codes of already installed applications relevant for this deployment from $env:COMPUTERNAME ..."

    Get-WmiObject Win32_Product | where { $_.Name -like "RSA.ESB.*" } | % {

        $entry = ($_.Name.Replace($_.Version,$null)).Trim()

        $installedMap.Set_Item($entry, $_)

    }

}

 

function GetProductCode($packageName, $version) {

    if([System.String]::IsNullOrEmpty($packageName)) { return $null }

    $hit = $installedMap[$packageName]

    if($hit -eq $null) { return $null }

    if($hit.Version -eq $version) {

        return $hit.IdentifyingNumber

    }

    return $null

}

 

function Uninstall($packageName, $version) {

    $productCode = GetProductCode $packageName $version

    return Uninstall($packageName, $version, $productCode)

}

 

 

function Uninstall($packageName, $version, $productCode, $REMOTE_LOG_FOLDER, $DataStamp, $btdfProjPath) { 

    $result=$true

    $prefix = "$($REMOTE_LOG_FOLDER)\$($DataStamp)"

    if(-not(Test-Path -Path $prefix)) { ni -ItemType Directory -Force -Path $prefix | Out-Null }

    $logFile = "$prefix\UnInstall{0}-{1}.log" -f $packageName,$DataStamp

    if($productCode -eq $null) { $productCode = "{}" }

    

 

    $mesg = "$($env:COMPUTERNAME): Uninstalling $version of $packageName using product code $productCode..."

    Write-Host $mesg

    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

 

    $MSIArguments = @(

        "/x"

        $productCode

        "/qn"

        "/norestart"

        "/l+*v"

        $logFile

    )

    $procHandle = Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow -PassThru -RedirectStandardOutput $logFile

    $mesg = "$($env:COMPUTERNAME): Completed MSI execution for $version of $packageName using product code $productCode..."

    Write-Host $mesg

    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

 

    if($procHandle.HasExited) {

        if($procHandle.ExitCode -gt 0) { $colour = "Red"; $result=$false } else { $colour = "Green" }

        $mesg = "$($env:COMPUTERNAME): Finished uninstalling $packageName with product code $productCode..." 

        Write-Host -ForegroundColor $colour $mesg

        $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        if($procHandle.ExitCode -gt 0) { 

            

            Get-WmiObject Win32_Product | where { $_.Name -like "RSA.ESB.*" } | % {

                $entry = ($_.Name.Replace($_.Version,$null)).Trim()

                if([System.String]::Equals($entry, $packageName, [System.StringComparison]::CurrentCultureIgnoreCase)) { $result = $false } else { $result=$true }

            }

            if(-not($result)) {

                $mesg = "$($env:COMPUTERNAME): ERROR*** Un-installation of $packageName with product code $productCode failed with exit code $($procHandle.ExitCode). Please check 'My Documents' folder for the error log $logFile to troubleshoot. ***NOTE - Please verify that you are a local ADMINISTRATOR on $($env:COMPUTERNAME)." 

                Write-host -ForegroundColor Red $mesg

                "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

            } else {

                $mesg = "$($env:COMPUTERNAME): WARNING: Verified that $packageName with product code $productCode was successfully UNinstalled inspite of a previous error. Please double check logs and state on the server." 

                Write-Host -f Yellow $mesg

                $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

            }            

        }

    } else {

        $result=$false

        $mesg = "$($env:COMPUTERNAME): Un-installation of $packageName with product code $productCode was triggered but completion status could not be ascertained. Please verify un-installation manually via Control Panel. If the package is still present please retry the script.  ***NOTE - Please verify that you are a local ADMINISTRATOR on $($env:COMPUTERNAME)."

        Write-Host -ForegroundColor Red $mesg

        "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

    }

    return $result

}

 

function DoParallelUninstall($application, $version, $servers) {

    

    #todo overload the function to accept existing $session. Can uninstall & deploy false in the same session

    $session = New-PSSession -ComputerName $servers -EnableNetworkAccess

            $result = $true

    

        $productCode = GetProductCode $application $version

        if($productCode -eq $null) { 

            $hit = $installedMap[$application]

            if($hit -ne $null) {

                $prevVersion = $version

                $version = $hit.Version

                $productCode = $hit.IdentifyingNumber 

                LogInfo "Version $prevVersion of $application was not found. Uninstalling the previous version $version instead..."

            }

        } #uninstall an earlier version if ver not present

        if($productCode -eq $null) {

            LogInfo "$version of $application was not found on server $($env:COMPUTERNAME). Skipping uninstallation..."

        } else {

            LogInfo "Inferred product code $productCode for $application version $version"

        }

        $appRootPath = [System.IO.Path]::Combine($prgFilesLoc, $application, $PROJECT_VERSION)

        $btdfProjPath = [System.IO.Path]::Combine($appRootPath, "Deployment", "Deployment.btdfproj")

 

        $allMSIFunctionDefs = "function Uninstall { ${function:Uninstall} };"

        $MSIfuncCallBlock = {

        Param( $allMSIFunctionDefs, $application, $version, $productCode, $REMOTE_LOG_FOLDER, $DataStamp, $btdfProjPath )

        . ([ScriptBlock]::Create($allMSIFunctionDefs))

        Uninstall $application $version $productCode $REMOTE_LOG_FOLDER $DataStamp $btdfProjPath

 

        }

        if($productCode -ne $null) {

            LogInfo "Uninstalling $application V$version on servers $servers"

            $opResults = Invoke-Command -Session $session -ScriptBlock $MSIfuncCallBlock -ArgumentList $allMSIFunctionDefs,$application, $version, $productCode, $REMOTE_LOG_FOLDER, $DataStamp, $btdfProjPath 

            

                                    $opResults | % {

                                               if($_ -eq $false) { 

                                                           $result = $false

                    LogError "Failed to Uninstall $application V$version on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern UnInstall$application*.log" $null

                    LogInfo "Failed to Uninstall $application V$version on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern UnInstall$application*.log" $null

 

                } else { 

                    LogInfo "Successfully Uninstalled $application V$version on server $($_.PSComputerName)"

                } 

                                    }

        }

    Remove-PSSession $session

            return $result

}

 

function Undeploy($application) {

    $appRootPath = [System.IO.Path]::Combine($prgFilesLoc, $application, $PROJECT_VERSION)

    $btdfProjPath = [System.IO.Path]::Combine($appRootPath, "Deployment", "Deployment.btdfproj")

    $commandPath = """$msBuildPath"""

    $msBuildCmdLineParams = "/p:Configuration=Server;DeployPDBsToGac=false /t:SetToolsVersionParam;ExportSettings;Undeploy ""$btdfProjPath"" /tv:4.0"

    $cmdLineParams = $MSBUILD_PARAMS + " " + $msBuildCmdLineParams

    $logFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\$($application)_Undeploy.log" 

    $errLogFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\$($application)_Undeploy_Errors.log" 

    $result = $true

 

    LogInfo "Started Undeploying application $application..."

    $procHandle = Start-Process -FilePath $commandPath -ArgumentList $cmdLineParams -Wait -NoNewWindow -PassThru -RedirectStandardOutput $logFile -RedirectStandardError $errLogFile 

    gc $logFile | Out-File -FilePath $runLogFile -Append -Encoding utf8 

    gc $errLogFile | Out-File -FilePath $runErrLogFile -Append -Encoding utf8     

    

    if($procHandle.HasExited) {

                        if(-not(gc $logFile | where { $_ -match "0 Error\(s\)" })) { $result = $false; }

        $errored=($procHandle.ExitCode -gt 0) -or -not($result)

        if($errored) { 

            if(($catalog.Applications | where { $_.Name -eq $application }) -eq $null) { $errored=$false; $result=$true }            

        } 

        if($errored) { $colour = "Red"; $result=$false }

        else { $colour = "Green" }

        $mesg = "Finished Undeploying $application..." 

        Write-Host -ForegroundColor $colour $mesg

        $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        

        if($errored) { 

            $mesg = "ERROR*** Undeploying $application failed with exit code $($procHandle.ExitCode). Please check $errLogFile and $runErrLogFile for the error log to troubleshoot." 

            Write-host -ForegroundColor Red $mesg

            "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        }

    } else {

        $mesg = "Undeploying $application was triggered but completion status could not be ascertained. Please verify manually via BizTalk Server Administration console. If the application is still present please terminate suspended/in progress service instances and retry the script. If this still fails please manually undeploy via BTDF and run the script thereafer."

        Write-Host -ForegroundColor Red $mesg

        "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

    }

    return $result

}

 

#todo test this by deploying MSI via deploy.bat & then redeploying

function Redeploy($application) {

    $appRootPath = [System.IO.Path]::Combine($prgFilesLoc, $application, $PROJECT_VERSION)

    $btdfProjPath = [System.IO.Path]::Combine($appRootPath, "Deployment", "Deployment.btdfproj")

    $commandPath = """$msBuildPath"""

    $msBuildCmdLineParams = "/p:Configuration=Server;DeployPDBsToGac=false /t:SetToolsVersionParam;InitSettingsFilePath;deploy ""$btdfProjPath"" /tv:4.0"

    $cmdLineParams = $MSBUILD_PARAMS + " " + $msBuildCmdLineParams

    $logFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\$($application)_Redeploy.log" 

    $errLogFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\$($application)_Redeploy_Errors.log" 

    $result = $true

 

    LogInfo "Started Redeploying application $application..."

    $procHandle = Start-Process -FilePath $commandPath -ArgumentList $cmdLineParams -Wait -NoNewWindow -PassThru -RedirectStandardOutput $logFile -RedirectStandardError $errLogFile

    gc $logFile | Out-File -FilePath $runLogFile -Append -Encoding utf8 

    gc $errLogFile | Out-File -FilePath $runErrLogFile -Append -Encoding utf8 

    

    if($procHandle.HasExited) {

                        if(-not(gc $logFile | where { $_ -match "0 Error\(s\)" })) { $result = $false; }

        if($procHandle.ExitCode -gt 0) { $colour = "Red"; $result=$false } else { $colour = "Green" }

        $mesg = "Finished Redeploying $application..." 

        Write-Host -ForegroundColor $colour $mesg

        $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        if($procHandle.ExitCode -gt 0 -or -not($result)) {

            $mesg = "ERROR*** Redeploying $application failed with exit code $($procHandle.ExitCode). Please check $errLogFile and $runErrLogFile for the error log to troubleshoot." 

            Write-host -ForegroundColor Red $mesg

            "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        }

    } else {

        $mesg = "Redeploying $application was triggered but completion status could not be ascertained. Please verify manually via BizTalk Server Administration console. If the application is NOT present please manually Redeploy via BTDF and run the script thereafer."

        Write-Host -ForegroundColor Red $mesg

        "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

    }

    return $result

}

 

# Done ONLY ON LOCAL server

function DeployTrue($application) {

    $srcFolder = [System.IO.Path]::Combine($msiSourcesRootFolder, $application)

    $batPath = [System.IO.Path]::Combine($srcFolder, $DEPLOY_BAT)

    $logFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\$($application)_Deploy.log" 

    $errLogFile = "$($LOCAL_LOG_FOLDER)\$($DataStamp)\$($application)_Deploy_Errors.log" 

    $result = $true

 

    LogInfo "Started Deploying application $application..."

    $procHandle = Start-Process -FilePath $batPath -ArgumentList "true S" -Wait -NoNewWindow -PassThru -RedirectStandardOutput $logFile -RedirectStandardError $errLogFile -WorkingDirectory $srcFolder

    gc $logFile | Out-File -FilePath $runLogFile -Append -Encoding utf8 

    gc $errLogFile | Out-File -FilePath $runErrLogFile -Append -Encoding utf8 

            

    if($procHandle.HasExited) {

                        if(-not(gc $logFile | where { $_ -match "0 Error\(s\)" })) { $result = $false; }

        if($procHandle.ExitCode -gt 0) { $colour = "Red"; $result=$false } else { $colour = "Green" }

        $mesg = "Finished Deploying $application..." 

        Write-Host -ForegroundColor $colour $mesg

        $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        if($procHandle.ExitCode -gt 0 -or -not($result)) { 

            $mesg = "ERROR*** Deploying $application failed with exit code $($procHandle.ExitCode). Please check $errLogFile and $runErrLogFile for the error log to troubleshoot." 

            Write-host -ForegroundColor Red $mesg

            "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        }

    } else {

        $mesg = "Deploying $application was triggered but completion status could not be ascertained. Please verify manually via BizTalk Server Administration console. If the application is NOT present please troubleshoot the situation by reviewing the error log file $errLogFile, log files $logFile and $runLogFile. Please run this script again after the problem has been addressed."

        Write-Host -ForegroundColor Red $mesg

        "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

    }

 

    #SSO Installation

    try{

    $SSOApps = gci "C:\Program Files (x86)\$application\1.0\*.app.xml" -Name

    Write-Host -ForegroundColor Red $SSOApps

    if($SSOApps.Length -eq 0){"No SSO applications to deploy"}

    else{"Following SSO application files found at C:\Program Files (x86)\$application\1.0"

        $SSOApps

        $SSOApps | foreach{

            [string]$SSOName = $_ -split '.app.xml'

            $SSOName = $SSOName.trim()

            

            if($Backup){SSO-BackupApp -SSOApplicationName $SSOName}

            SSO-InstallApp -SSOApplicationName $SSOName $application 

            }

        }

        }

        catch{

         Write-Host -ForegroundColor Green "SSO Not Imported"

        }

    return $result

}

 

 

 

function Install($packageName, $version, $REMOTE_MSI_FOLDER, $REMOTE_LOG_FOLDER, $DataStamp) { 

    $result=$true

    $prefix = "$($REMOTE_LOG_FOLDER)\$($DataStamp)"

    if(-not(Test-Path -Path $prefix)) { ni -ItemType Directory -Force -Path $prefix | Out-Null }

    $logFile = "$prefix\Install{0}-{1}.log" -f $packageName,$DataStamp

    $msiName = "$($packageName)-$($version).msi"

    $msiPath = [System.IO.Path]::Combine($REMOTE_MSI_FOLDER, $packageName, $msiName)

    

    $mesg = "$($env:COMPUTERNAME): Installing $version of $packageName using the MSI $msiPath..."

    Write-Host $mesg

    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

 

    $MSIArguments = @(

        "/i"

        $msiPath

        "/qn"

        "/norestart"

        "/l+*v"

        $logFile

    )

    $procHandle = Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow -PassThru -RedirectStandardOutput $logFile

    $mesg = "$($env:COMPUTERNAME): Completed execution of $version of $packageName using the MSI $msiPath..."

    Write-Host $mesg

    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

 

    if($procHandle.HasExited) {

        if($procHandle.ExitCode -gt 0) { $colour = "Red"; $result=$false } else { $colour = "Green" }

        $mesg = "$($env:COMPUTERNAME): Finished installing $msiName version $version..." 

        Write-Host -ForegroundColor $colour $mesg

        $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        if($procHandle.ExitCode -gt 0) { 

            if(Test-Path -Path "$($msiPath)") {

                Get-WmiObject Win32_Product | where { $_.Name -like "RSA.ESB.*" } | % {

                    $entry = ($_.Name.Replace($_.Version,$null)).Trim()

                    if([System.String]::Equals($entry, $packageName, [System.StringComparison]::CurrentCultureIgnoreCase)) { $result = $true } else { $result=$false }

                }

                if(-not($result)) {

                    $mesg = "$($env:COMPUTERNAME): ERROR*** Installation failed with exit code $($procHandle.ExitCode) using the MSI $msiPath. Please check 'My Documents' folder for the error log $logFile to troubleshoot." 

                    Write-host -ForegroundColor Red $mesg

                    "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                } else {

                    $mesg = "$($env:COMPUTERNAME): WARNING: Verified that $packageName with product code $productCode was successfully Installed inspite of a previous error. Please double check logs and state on the server." 

                    Write-Host -f Yellow $mesg

                    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                }

 

                

            } else {

                $mesg = "$($env:COMPUTERNAME): ERROR*** '$($msiPath)' could not be found. Please verify that you are a local ADMINISTRATOR on all servers in the BizTalk group, the MSI exists and is accessible."

                Write-host -ForegroundColor Red $mesg

                $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

            }

        }

    } else {

        $result=$false

        $mesg = "$($env:COMPUTERNAME): Installation of $msiName version $version was triggered but completion status could not be ascertained. Please verify installation via Control Panel or manually install the MSI $msiPath."

        Write-Host -ForegroundColor Red $mesg

        "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

    }

    return $result

}

 

#function DoParallelInstall($application, $version, $servers) {

    #todo overload the function to accept existing $session. Can uninstall & deploy false in the same session

 #   $session = New-PSSession -ComputerName $servers -EnableNetworkAccess

  #          $result = $true

 

   # $allMSIFunctionDefs = "function Install { ${function:Install} };"

#    $MSIfuncCallBlock = {

#    Param( $allMSIFunctionDefs, $application, $version, $REMOTE_MSI_FOLDER, $REMOTE_LOG_FOLDER, $DataStamp )

 #   . ([ScriptBlock]::Create($allMSIFunctionDefs))

  #  Install $application $version $REMOTE_MSI_FOLDER $REMOTE_LOG_FOLDER $DataStamp

 

   # }

    

    #LogInfo "Installing $application V$version on servers $servers"

    #$opResults = Invoke-Command -Session $session -ScriptBlock $MSIfuncCallBlock -ArgumentList $allMSIFunctionDefs, $application, $version, $REMOTE_MSI_FOLDER , $REMOTE_LOG_FOLDER, $DataStamp

    

     #       $opResults | % {

      #                  if($_ -eq $false) { 

       #                             $result = $false

        #    LogError "Failed to Install $application V$version on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern Install$application*.log" $null

         #   LogInfo "Failed to Install $application V$version on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern Install$application*.log" $null

 

        #} else { 

         #   LogInfo "Successfully Installed $application V$version on server $($_.PSComputerName)"

        #} 

         #   }

            

  #  Remove-PSSession $session

 #           return $result

#}

 

function QuickDeploy($application, $btdfProjPath, $commandPath, $REMOTE_LOG_FOLDER, $DataStamp) {

    $msBuildCmdLineParams = "/p:Configuration=Server;IncludeSSO=false;DeployPDBsToGac=false /t:SetToolsVersionParam;UpdateOrchestration ""$btdfProjPath"" /tv:4.0"

    $cmdLineParams = $msBuildCmdLineParams

    $prefix = "$($REMOTE_LOG_FOLDER)\$($DataStamp)"

    if(-not(Test-Path -Path $prefix)) { ni -ItemType Directory -Force -Path $prefix | Out-Null }

    $logFile = "$($prefix)\QuickDeploy$($application)-$($DataStamp).log" 

    $errLogFile = "$($prefix)\QuickDeploy$($application)-$($DataStamp)_Errors.log" 

    $result = $true

 

    $mesg = "$($env:COMPUTERNAME): Started QuickDeploying application $application..."

    Write-Host $mesg

    

    $procHandle = Start-Process -FilePath $commandPath -ArgumentList $cmdLineParams -Wait -NoNewWindow -PassThru -RedirectStandardOutput $logFile -RedirectStandardError $errLogFile 

    

    if($procHandle.HasExited) {

                        if(-not(gc $logFile | where { $_ -match "0 Error\(s\)" })) { $result = $false; }

        if($procHandle.ExitCode -gt 0) { $colour = "Red"; $result=$false } else { $colour = "Green" }

        $mesg = "$($env:COMPUTERNAME): Finished QuickDeploying $application..." 

        Write-Host -ForegroundColor $colour $mesg

        $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        if($procHandle.ExitCode -gt 0 -or -not($result)) { 

            if(Test-Path -Path "$($btdfProjPath)") {

                $mesg = "$($env:COMPUTERNAME): ERROR*** QuickDeploying $application failed with exit code $($procHandle.ExitCode). Please check $errLogFile and $logFile for the error log to troubleshoot." 

                Write-host -ForegroundColor Red $mesg

                "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                

            } else {

                $mesg = "$($env:COMPUTERNAME): ERROR*** '$($btdfProjPath)' could not be found. Please verify that you are a local ADMINISTRATOR on all servers in the BizTalk group, '$application' is installed and the BTDF project file exists."

                Write-host -ForegroundColor Red $mesg

                $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

            }

        }

    } else {

        $mesg = "$($env:COMPUTERNAME): QuickDeploying $application was triggered but completion status could not be ascertained. Please verify manually via BizTalk Server Administration console. If the application is still present please terminate suspended/in progress service instances and retry the script. If this still fails please manually delete all entries associated with $application from GAC on server $($env:COMPUTERNAME) and run the script thereafer."

        Write-Host -ForegroundColor Red $mesg

        "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

    }

    return $result

}

 

function DoQuickDeploy($application, $servers) {

    #todo overload the function to accept existing $session. Can uninstall & deploy false in the same session 

    $appRootPath = [System.IO.Path]::Combine($prgFilesLoc, $application, $PROJECT_VERSION)

    $btdfProjPathQD = [System.IO.Path]::Combine($appRootPath, "Deployment", "Deployment.btdfproj")

    $commandPath = """$msBuildPath"""

    $session = New-PSSession -ComputerName $servers -EnableNetworkAccess

            $result = $true

 

    $quickDeployFunctionDefs = "function QuickDeploy { ${function:QuickDeploy} };"

    $quickDeployfuncCallBlock = {

    Param( $quickDeployFunctionDefs, $application, $btdfProjPath, $commandPath, $REMOTE_LOG_FOLDER, $DataStamp )

    . ([ScriptBlock]::Create($quickDeployFunctionDefs))

    QuickDeploy $application $btdfProjPath $commandPath $REMOTE_LOG_FOLDER $DataStamp

 

    }

    

    LogInfo "Quick deploying $application on servers $servers"

    $opResults = Invoke-Command -Session $session -ScriptBlock $quickDeployfuncCallBlock -ArgumentList $quickDeployFunctionDefs, $application, $btdfProjPathQD, $commandPath, $REMOTE_LOG_FOLDER, $DataStamp

    

    $opResults | % {

        if($_ -eq $false) { 

                                    $result = $false

            LogError "Failed to QuickDeploy $application on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern QuickDeploy$application*.log" $null

            LogInfo "Failed to QuickDeploy $application on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern QuickDeploy$application*.log" $null

 

        } else { 

            LogInfo "Successfully QuickDeployed $application on server $($_.PSComputerName)"

        } 

    }

 

    Remove-PSSession $session

            return $result

}

 

function WebDeploy($packageName){

 

$FacadeFolder = "C:\Program Files (x86)\$packageName\1.0\Facade"

             $Manifest = "C:\windows\temp\$packageName\Manifest\FacadeManifest.xml"

 

#$PackageFolder = "D:\Nolio-Deploy\TST_2016-07-11_16_12_40_002\Packages\RSA.ESB.Services.Document.1.0.0-COE1054"

#$FacadeFolder = "$PackageFolder\Facade"

#$Manifest = "$PackageFolder\Manifest\FacadeManifest.xml"

 

 

import-Module WebAdministration

 

Function Get-SitePath{

        $sitePath = ""

        }

 

Function Get-AppPool{

        $appPool = (get-childItem IIS:\appPools | Where{$_.Name -eq $appPoolName})

        if(-Not $appPool){"AppPool $appPoolName not found on local IIS";exit 1}

        else{$script:LocalAppPool = $appPool

             $script:AppPoolState = ($appPool.State)}

        }

 

Function Backup-WebApp{

        try{

        copy-Item $webAppPath $pwd\Backup\$webAppName -Recurse -Force

        "ServiceFacade backed up to $pwd\Backup\$webAppName"}

        catch{"Error trying to backup"

              Write-Warning -Message $_.Exception.Message ;exit 1}

        }

            

Function PreProcess-WebConfig{

 

    $xmlPreprocArgs = "/v /q /c /nologo /i  `"$FacadeFolder\Web.Config`" $PreProcSettings /e $Environment"

    

    $exe = (Start-Process -FilePath .\XmlPreprocess.exe -ArgumentList $xmlPreprocArgs -PassThru -Wait -NoNewWindow)

    if ($exe.exitcode -eq 0){"Web.Config Processing completed"}

    else{"Error in xmlpreprocessing of $WebAppName web.config.";exit 1}

    }

 

Function PreProcess-FacadeManifest{

 

    $xmlPreprocArgs = "/v /q /c /nologo /i  `"$Manifest`" $PreProcSettings /e $Environment"

    

    $exe = (Start-Process -FilePath .\XmlPreprocess.exe -ArgumentList $xmlPreprocArgs -PassThru -Wait -NoNewWindow)

    if ($exe.exitcode -eq 0){"Facade manifest Processing completed"}

    else{"Error in xmlpreprocessing of $WebAppName facadeManifest.";exit 1}

    }

 

Function Copy-WebContent{

    try{

    #Copy-Item $FacadeFolder\* $webAppPath -Recurse -Force

    xcopy $FacadeFolder $webAppPath /E /I

    "content copied to $webAppPath"}

    catch{"Error trying to copy content"

          Write-Warning -Message $_.Exception.Message ;exit 1}

    }

 

Function Delete-WebContent{

    try{

    Remove-Item $WebAppPath\* -Recurse -Force}

    catch{"error deleting content"

          Write-Warning -Message $_.Exception.Message ;exit 1}

    }

 

Function Create-Webapp{

    try{New-WebApplication -Name $WebAppName  -Site $SiteName -PhysicalPath $webAppPath -ApplicationPool $AppPoolName}

    catch{"error Creating webApplication"

          Write-Warning -Message $_.Exception.Message ;exit 1}

    }

 

Function Stop-AppPool{

    Stop-WebAppPool $AppPoolName

    do{Get-AppPool;"Stop Pending"}

    until ($AppPoolState -eq "Stopped")

    }

 

Function Start-AppPool{

    Start-WebAppPool $AppPoolName

    do{Get-AppPool;"Start Pending"}

    until ($AppPoolState -eq "Started")

    }

 

Write-Host "working directory $pwd"

 

Write-Host "In Web deploy $FacadeFolder"

 

                 if(test-path $Manifest){

 

#PreProcess-FacadeManifest

 

$XpathSiteName = 'FacadeManifest/WebSite/text()'

$XpathAppName = 'FacadeManifest/ApplicationName/text()'

$XpathAppPool = 'FacadeManifest/AppPool/text()'

 

Write-Host "In Web deploy facade  $XpathSiteName"

 

 

$WebAppName = Select-Xml -Path $Manifest -XPath $XpathAppName

$WebAppName = $WebAppName.Node.Value

$AppPoolName = Select-Xml -Path $Manifest -XPath $XpathAppPool

$AppPoolName = $AppPoolName.Node.Value

$SiteName = Select-Xml -Path $Manifest -XPath $XpathSiteName

$SiteName = $SiteName.Node.Value

 

Write-Host "In Web deploy $SiteName"

if(-not $SiteName){"Error! no value found for SiteName";exit 1}

else{"Website is $SiteName"}

if(-not $WebAppName){"Error! no value found for WebApplication name";exit 1}

else{"webApp is $WebAppName"}

if(-not $AppPoolName){"Error! no value found for ApplicationPool";exit 1}

else{"AppPool is $AppPoolName"}

 

 

#Get-SitePath

$webAppPath = "D:\IISROOT\iis\inetpub\wwwroot\$webAppName"

 

 

if (-Not(Test-Path $webAppPath)){"$webAppPath not present. Nothing to backup"}

else{"$WebAppName found. Proceeding with backup"

    Backup-WebApp

    Delete-WebContent}

 

#PreProcess-WebConfig

 

Get-AppPool

if ($AppPoolState -eq "Started"){"$appPoolName is running. Stopping to update content";Stop-AppPool}

else {"$appPoolName is in $AppPoolState state"}

 

Copy-WebContent

 

 

#GCI on Site return mixed typename if vDir are present

#$WebApplication = Get-ChildItem IIS:\Sites\$SiteName | Where {$_.Name -eq $WebAppName -and $_.NodeType -eq "application"}

$WebApplication = Get-WebApplication | Add-Member -MemberType ScriptProperty -Name ApplicationName -Value { $this.Path.Trim('/') } -PassThru | Where {$_.ApplicationName -eq $WebAppName}

if (-Not $WebApplication){"$WebAppName not found. Creating and assigning to $AppPoolName"

    Create-Webapp}

 

Get-AppPool

if ($AppPoolState -eq "Stopped"){"$appPoolName is stopped. Starting!";Start-AppPool}

else {"$appPoolName is in $AppPoolState state"}

 

 

}

else{"No servicefacade to deploy"}

 

 

}

 

 

function DoWebDeploy($application, $servers) {

    #todo overload the function to accept existing $session. Can uninstall & deploy false in the same session

   

    $session = New-PSSession -ComputerName $servers -EnableNetworkAccess

 

    $quickDeployFunctionDefs = "function WebDeploy { ${function:WebDeploy} };"

    $quickDeployfuncCallBlock = {

    Param( $quickDeployFunctionDefs, $application)

    . ([ScriptBlock]::Create($quickDeployFunctionDefs))

    WebDeploy $application

 

    }

    

    LogInfo "Web deploying $application on servers $servers using settings from '$($settingsFile)'"

    $opResults = Invoke-Command -Session $session -ScriptBlock $quickDeployfuncCallBlock -ArgumentList $quickDeployFunctionDefs, $application

    LogInfo "Web deploying $opResults $application"

            $opResults | % {

                        if($_ -eq $false) { 

            LogError "Failed to Web Deploy $application on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern IISQuickDeploy$application*.log" $null

            LogInfo "Failed to Web Deploy $application on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern IISQuickDeploy$application*.log" $null

 

        } else { 

            LogInfo "Successfully Deployed Web $application on server $($_.PSComputerName)"

        } 

            }

            

    Remove-PSSession $session

            # for IISQuickDeploy we always report success. Else execution will stop on error. The assumption is that failure on 1 web server should not stop further deployment steps.

            return $true

}

 

function IISQuickDeploy($application, $settingsFile, $btdfProjPath, $commandPath, $REMOTE_LOG_FOLDER, $DataStamp) {

            $msBuildCmdLineParams = """$btdfProjPath"" /p:Configuration=Server;PropsFromEnvSettings=true;IncludeSSO=true;DeployPDBsToGac=false;IncludeVirtualDirectories=true;ENV_SETTINGS=""$settingsFile"" /t:SetPropertiesFromEnvironmentSettings;DeployVDirs;CustomPostDeployTarget"

 

    $cmdLineParams = $msBuildCmdLineParams

    $prefix = "$($REMOTE_LOG_FOLDER)\$($DataStamp)"

    if(-not(Test-Path -Path $prefix)) { ni -ItemType Directory -Force -Path $prefix | Out-Null }

    $logFile = "$($prefix)\IISQuickDeploy$($application)-$($DataStamp).log" 

    $errLogFile = "$($prefix)\IISQuickDeploy$($application)-$($DataStamp)_Errors.log" 

    $result = $true

 

    $mesg = "$($env:COMPUTERNAME): Started IISQuickDeploying application $application..."

    Write-Host $mesg

    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

 

            

            $vDirList = $null

            # First cleanup existing VDirs

            if(-not(Get-Module -Name WebAdministration)) { try { Import-Module WebAdministration } catch {} }

    if(Get-Module -Name WebAdministration) {

        if(Test-Path -Path "$($btdfProjPath)") {

            $xml = [xml](Get-Content "$($btdfProjPath)")

            if($xml -ne $null) {

                $vDirList = $xml.Project.ItemGroup.VDirList

                                               if($vDirList -ne $null) {

                                                           $vDirList | % {

                                                                       $vdir = $_.Vdir                        

                        $vdirPropertyTokens =  ([regex]"(\$\((.+)\))").Match($vdir)

                                                                       if($vdirPropertyTokens.Groups -and $vdirPropertyTokens.Groups.Count -gt 2) {

                                                                                   $vdirPropertyToken = $vdirPropertyTokens.Groups[2].Value

                            $tempVdirSubstitution = $xml.Project.PropertyGroup.$vdirPropertyToken

                            if($tempVdirSubstitution -and $tempVdirSubstitution.Length -gt 0) { 

                                $vdir = $tempVdirSubstitution[0] 

                            } 

                                                                       }

                                                                       $webApp = $null

                                                                       if(-not([System.String]::IsNullOrEmpty($vdir))) {

                                                                                   $webApp = Get-WebApplication -Site "Default Web Site" -Name $vdir

                                                                                   try {

                                                                                               if($webApp) {

                                                                                                          $mesg = "$($env:COMPUTERNAME): Removing existing Web Application $($vdir)..."

                                                                                                          Write-Host $mesg

                                                                                                          $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                                          remove-WebApplication -Site "Default Web Site" -Name $vdir

                                                                                                          $mesg = "$($env:COMPUTERNAME): Successfully removed the existing Web Application $($vdir)..."

                                                                                                          Write-Host $mesg

                                                                                                          $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                                          

                                                                                                          try {

                                                                                                                      $appPool = $webApp.applicationPool

                                                                                                                      try {

                                                                                                                                  $appPoolState = Get-WebAppPoolState -Name $appPool

                                                                                                                                  if($appPoolState -ne $null -and $appPoolState.Value -ne "Stopped") {

                                                                                                                                              Stop-WebAppPool -ErrorAction Ignore -Name $appPool

                                                                                                                                              $mesg = "$($env:COMPUTERNAME): Stopped AppPool $appPool..."

                                                                                                                                              Write-Host $mesg

                                                                                                                                              $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                                                                  }

                                                                                                                      } catch {}

                                                                                                          } catch {}

                                                                                               }                                                                      

                                                                                   } catch [System.Exception] { 

                                                                                               $excp = $_.Exception

                                                                                               if(Get-WebApplication -Site "Default Web Site" -Name $vdir) {

                                                                                                          try {

                                                                                                                      Sleep 5

                                                                                                                      Remove-WebApplication -Site "Default Web Site" -Name $vdir

                                                                                                          }

                                                                                                          catch {}

                                                                                                          if(Get-WebApplication -Site "Default Web Site" -Name $vdir) {

                                                                                                                      $mesg = "$($env:COMPUTERNAME): WARNING: Failed to remove the existing IIS web application $($vdir). If Installation of the new version of $($vdir) fails at a later stage please manually remove $($vdir) using IIS Manager console and re-run deploy.ps1. Details - $excp" 

                                                                                                                      Write-Host -f Yellow  $mesg 

                                                                                                                      $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                                          }

                                                                                               } else {

                                                                                                          $mesg = "$($env:COMPUTERNAME): Successfully removed the existing Web Application $($vdir)..."

                                                                                                          Write-Host $mesg

                                                                                                          $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                               }

                                                                                   }

                                                                                   # Now remove IIS metabase configuration

                                                                                   $appHostPath="'MACHINE/WEBROOT/APPHOST'"

                                                                                   $vdirConfig="Default Web Site/$vdir"

                                                                                   

                                                                                   $removalAttempts = 0

                                                                                   $maxRemovalAttempts = 5

                                                                                   $removalResult = $false

                                                                                   while(-not($removalResult) -and ($removalAttempts -lt $maxRemovalAttempts))

                                                                                   {

                                                                                               $removalResult = $true

                                                                                               try {

                                                                                                          if(Get-WebConfigurationLocation -Name "$($vdirConfig)") {                                                                                 

                                                                                                                      $mesg = "$($env:COMPUTERNAME): Attempt $($removalAttempts+1) to Remove existing IIS metabase location $($vdirConfig) ..."

                                                                                                                      Write-Host $mesg

                                                                                                                      $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                                                      Remove-WebConfigurationLocation -Name "$($vdirConfig)"

                                                                                                                      $mesg = "$($env:COMPUTERNAME): Successfully removed the existing existing IIS metabase location $($vdirConfig)..."

                                                                                                                      Write-Host $mesg

                                                                                                                      $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                                          }

                                                                                               }

                                                                                               catch [System.Exception]{

                                                                                                          $removalResult = $false

                                                                                                          $excp = $_.Exception

                                                                                                          try {

                                                                                                                      if($webApp) {

                                                                                                                                  $appPool = $webApp.applicationPool

                                                                                                                                  try {

                                                                                                                                              $appPoolState = Get-WebAppPoolState -Name $appPool

                                                                                                                                              if($appPoolState -ne $null -and $appPoolState.Value -ne "Stopped") {

                                                                                                                                                          Stop-WebAppPool -ErrorAction Ignore -Name $appPool

                                                                                                                                              }

                                                                                                                                  } catch {}

                                                                                                                      }

                                                                                                          } catch {}

                                                                                               }

                                                                                               Sleep 5

                                                                                               $found = Get-WebConfigurationLocation -Name "$($vdirConfig)"

                                                                                               if($found -and ($removalAttempts -ge $maxRemovalAttempts) ) {                                                                         

                                                                                                          $mesg = "$($env:COMPUTERNAME): WARNING: Failed to remove the existing IIS metabase location $($vdirConfig). If Installation of the new version of $($vdir) fails at a later stage please manually remove the location $($vdirConfig) IIS metabase on server $($env:COMPUTERNAME) and re-run deploy.ps1. Details - $excp" 

                                                                                                          Write-Host -f Yellow  $mesg 

                                                                                                          $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                               } elseif(-not($found))  {

                                                                                                          $removalResult = $true

                                                                                                          $mesg = "$($env:COMPUTERNAME): Successfully removed the existing existing IIS metabase location $($vdirConfig)..."

                                                                                                          Write-Host $mesg

                                                                                                          $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                               }           

                                                                                   }                                                                                  

                                                                       }

                                                           }

                                               }

                                    }

                        }

            }

            $mesg = "$($env:COMPUTERNAME): Starting IIS deployment of $application..." 

            Write-Host -ForegroundColor "White" $mesg

            $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

    

            $result = $false

            $attempt = 0

            $attemptMax = 7

    $msbTempOutBuffer = "temp.out"

    $msbTempErrBuffer = "temp.err"

            

            while(-not($result) -and ($attempt -lt $attemptMax)) {         

                        $result = $true

                        $procHandle = Start-Process -FilePath $commandPath -ArgumentList $cmdLineParams -Wait -NoNewWindow -PassThru -RedirectStandardOutput $msbTempOutBuffer -RedirectStandardError $msbTempErrBuffer

                        gc $msbTempOutBuffer | Out-File -FilePath $logFile -Append -Encoding utf8  

                        gc $msbTempErrBuffer | Out-File -FilePath $errLogFile -Append -Encoding utf8  

                        

                        if($procHandle.HasExited) {

                                    if($procHandle.ExitCode -gt 0) { $colour = "Red"; $result=$false } else { $colour = "White" }

                                    $mesg = "$($env:COMPUTERNAME): Executed IISQuickDeploy of $application..." 

                                    Write-Host -ForegroundColor $colour $mesg

                                    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                                    if($procHandle.ExitCode -gt 0) { 

                                               $result = $false

                                               if(-not(Test-Path -Path "$($btdfProjPath)")) {   

                                                           $mesg = "$($env:COMPUTERNAME): ERROR*** '$($btdfProjPath)' could not be found. Please verify that you are a local ADMINISTRATOR on all servers in the BizTalk group, '$application' is installed and the BTDF project file exists."

                                                           Write-host -ForegroundColor Red $mesg

                                                           $mesg | Out-File -FilePath $logFile -Append -Encoding utf8 

                                                           return $false                                                                       

                                               }                                   

                                    }

                        } elseif($attempt -ge $attemptMax) {

                                    $result = $false

                                    $mesg = "$($env:COMPUTERNAME): IISQuickDeploying $application was triggered but completion status could not be ascertained. Please verify manually via BizTalk Server Administration console. If the application is still present please terminate suspended/in progress service instances and retry the script. If this still fails please manually delete all entries associated with $application from GAC on server $($env:COMPUTERNAME) and run the script thereafer."

                                    Write-Host -ForegroundColor Red $mesg

                                    "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                                    return $result

                        } else {

                                    $result = $false

                                    Sleep 3                 

                        }

                        

                        $colour = "White"

                        if($vDirList -ne $null) {

                                    $mesg = "$($env:COMPUTERNAME): Verifying IIS deployment of $application..." 

                                    Write-Host -ForegroundColor $colour $mesg

                                    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                                                           

                                    $vDirList | where { $_ -ne $null } | % {

                                               $vdir = $_.Vdir

                                                $specifiedAppPool = $_.AppPool

                                                $vdirPropertyTokens =  ([regex]"(\$\((.+)\))").Match($vdir)

                                                if($vdirPropertyTokens.Groups -and $vdirPropertyTokens.Groups.Count -gt 2) {

                                                           $vdirPropertyToken = $vdirPropertyTokens.Groups[2].Value

                    $tempVdirSubstitution = $xml.Project.PropertyGroup.$vdirPropertyToken

                    if($tempVdirSubstitution -and $tempVdirSubstitution.Length -gt 0) { 

                        $vdir = $tempVdirSubstitution[0] 

                    } 

 

                                               }

                $poolPropertyTokens =  ([regex]"(\$\((.+)\))").Match($specifiedAppPool)

                if($poolPropertyTokens.Groups -and $poolPropertyTokens.Groups.Count -gt 2) {

                                                           $poolPropertyToken = $poolPropertyTokens.Groups[2].Value

                    $tempPoolSubstitution = $xml.Project.PropertyGroup.$poolPropertyToken

                    if($tempPoolSubstitution -and $tempPoolSubstitution.Length -gt 0) { 

                        $subPoolTokens = ([regex]"(\$\((.+)\))").Match($tempPoolSubstitution)

                        if($subPoolTokens.Groups -and $subPoolTokens.Groups.Count -gt 2) {

                            $poolSubPropertyToken = $subPoolTokens.Groups[2].Value

                            $poolSubPropertyTokenValue = $xml.Project.PropertyGroup.$poolSubPropertyToken

                            if($poolSubPropertyTokenValue -and $poolSubPropertyTokenValue.Length -gt 0) { 

                                $poolSubPropertyTokenValue = $poolSubPropertyTokenValue[0] 

                                $specifiedAppPool=([regex]"((.*)\$\((.+)\))(.*)").Replace($tempPoolSubstitution, "`$2$poolSubPropertyTokenValue`$4").Trim()

                            } 

                        }

                    } 

                                               }

                                               $mesg = "$($env:COMPUTERNAME): Verifying virtual directory $vdir..." 

                                               Write-Host -ForegroundColor $colour $mesg

                                               $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                                               if(-not([System.String]::IsNullOrEmpty($vdir))) {

                                                           $webApp = Get-WebApplication -Site "Default Web Site" -Name $vdir

                                                           try {

                                                                       if($webApp -eq $null) { $result = $false }

                                                                       else {

                                                                                   if($webApp.applicationPool -ne $specifiedAppPool) { 

                                                                                               $mesg = "$($env:COMPUTERNAME): App pool $specifiedAppPool was not correctly assigned for $vdir..." 

                                                                                               Write-Host -ForegroundColor $colour $mesg

                                                                                               $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                               $result = $false 

                                                                                   }

                                                                                   $path = Get-Item -Path $webApp.PhysicalPath -ErrorAction Ignore

                                                                       if([System.String]::IsNullOrEmpty($path.Parent.Parent.FullName)) { $result = $false }

                                                                                   elseif(-not($btdfProjPath.StartsWith($path.Parent.Parent.FullName))) { 

                                                                                               $mesg = "$($env:COMPUTERNAME): Virtual Dir path was not correctly assigned for $vdir..." 

                                                                                               Write-Host -ForegroundColor $colour $mesg

                                                                                               $mesg | Out-File -FilePath $logFile -Append -Encoding utf8

                                                                                               $result = $false 

                                                                                   }                                                                                              

                                                                       }

                                                           } catch {}

                                               }

                                    }

                        }

                        $attempt = $attempt + 1             

                        if(-not($result)) {                                                     

                                    $mesg = "$($env:COMPUTERNAME): IISQuickDeploy attemp# $($attempt+1) of $application..." 

                                    Write-Host -ForegroundColor "Yellow" $mesg

                                    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8 

                                    Sleep 4                                                     

                        }                                   

            }

            if(-not($result)) {

                        $mesg = "$($env:COMPUTERNAME): ERROR*** IISQuickDeploying $application failed. Please check $errLogFile and $logFile to troubleshoot." 

                        Write-host -ForegroundColor Red $mesg

                        "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

            } else {

                        $mesg = "$($env:COMPUTERNAME): IISQuickDeploy post execution verification was successful after $attempt attempt(s) for all virtual directories in $application..." 

                        Write-Host -ForegroundColor "Yellow" $mesg

                        $mesg | Out-File -FilePath $logFile -Append -Encoding utf8 

            }                       

    return $result

}

 

function DoIISQuickDeploy($application, $servers) {

    #todo overload the function to accept existing $session. Can uninstall & deploy false in the same session

    $appRootPath = [System.IO.Path]::Combine($prgFilesLoc, $application, $PROJECT_VERSION)

    $btdfProjPathQD = [System.IO.Path]::Combine($appRootPath, "Deployment", "Deployment.btdfproj")

    $commandPath = """$msBuildPath"""

    $settingsFile = $settingsMap[$application]

            $settingsFilePath = "$REMOTE_MSI_FOLDER\$application\$settingsFile"

    $session = New-PSSession -ComputerName $servers -EnableNetworkAccess

 

    $quickDeployFunctionDefs = "function IISQuickDeploy { ${function:IISQuickDeploy} };"

    $quickDeployfuncCallBlock = {

    Param( $quickDeployFunctionDefs, $application, $settingsFile, $btdfProjPath, $commandPath, $REMOTE_LOG_FOLDER, $DataStamp )

    . ([ScriptBlock]::Create($quickDeployFunctionDefs))

    IISQuickDeploy $application $settingsFile $btdfProjPath $commandPath $REMOTE_LOG_FOLDER $DataStamp

 

    }

    

    LogInfo "IISQuick deploying $application on servers $servers using settings from '$($settingsFile)'"

    $opResults = Invoke-Command -Session $session -ScriptBlock $quickDeployfuncCallBlock -ArgumentList $quickDeployFunctionDefs, $application, $settingsFilePath, $btdfProjPathQD, $commandPath, $REMOTE_LOG_FOLDER, $DataStamp

    

            $opResults | % {

                        if($_ -eq $false) { 

            LogError "Failed to IISQuickDeploy $application on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern IISQuickDeploy$application*.log" $null

            LogInfo "Failed to IISQuickDeploy $application on server $($_.PSComputerName). To troubleshoot, please see the log file in 'My Documents' folder with the name pattern IISQuickDeploy$application*.log" $null

 

        } else { 

            LogInfo "Successfully IISQuickDeployed $application on server $($_.PSComputerName)"

        } 

            }

            

    Remove-PSSession $session

            # for IISQuickDeploy we always report success. Else execution will stop on error. The assumption is that failure on 1 web server should not stop further deployment steps.

            return $true

}

 

function CopyMSIsAcrossServers($application, $servers) {

    if(-not([System.String]::IsNullOrWhiteSpace($application))) {

        $srcFolder = [System.IO.Path]::Combine($msiSourcesRootFolder, $application)

        $srcConfigFolder = [System.IO.Path]::Combine((Get-Item $srcFolder).Parent.Parent,'config')

        if(Test-Path -path $srcFolder) {

            LogInfo "Copying contents of $srcFolder to c$\windows\temp on servers $servers" 

            $servers | % { 

                Copy-Item -path $srcFolder -Destination \\$_\c$\windows\temp   -Recurse -Force

            }

 

         if(Test-Path -path $srcConfigFolder) {

            LogInfo "Copying contents of $srcFolder to c$\windows\temp on servers $servers" 

            $servers | % { 

                Copy-Item -path $srcConfigFolder -Destination \\$_\c$\windows\temp   -Recurse -Force

            }

            }

            $result = $true

            $servers | % { 

                $wasCopied = Test-Path -path \\$_\c$\windows\temp\$($application) 

                if(-not($wasCopied)) {

                    LogError "Failed to copy $application to server $_."

                    LogInfo "Terminating deployment as $application folder contents could not be copied to server $_. Please verify that the server is online and that you are an Administrator on all servers in the BizTalk group."

                    LogInfo "Please re-run deploy.ps1 after restoring access..."

                    

                    Exit

                }

                $result = $result -and $wasCopied

            }

            if($result) { 

                LogInfo "Copy completed successfully..." 

            }

            else {

                LogInfo "Terminating deployment as $application folder contents could not be copied to all servers. Please verify that each server is online and that you are an Administrator on all servers in the BizTalk group."

                LogInfo "Please re-run deploy.ps1 after restoring access..."

                Exit

            }

        }

    }

}

 

function CleanupMSIsAcrossServers($application, $servers) {

    if(-not([System.String]::IsNullOrWhiteSpace($application))) {

        LogInfo "Deleting contents of c$\windows\temp\$application on servers $servers" 

        $servers | % { 

            if(Test-Path -Path \\$_\c$\windows\temp\$application  ) {

                Remove-Item -Recurse -Force \\$_\c$\windows\temp\$application  

            }

        }    

    }

}

 

function CheckCUs() {

    $btsCUs = @()

    $hisCUs = @()

    LogInfo "Analyzing cumulative updates..."

    gci -path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -ErrorAction SilentlyContinue | 

    where { $_.GetValue("DisplayName") -match "CU"} | % { 

        $item = $_.GetValue("DisplayName")

        if($item -match "Microsoft BizTalk Server" ) {  $btsCUs = $btsCUs + $item }

        if($item -match "Host Integration Server" ) {  $hisCUs = $hisCUs + $item }        

    }

    LogInfo "The following BizTalk CUs are installed in the current BizTalk group."

    $btsCUs | % { LogInfo $_ }     

    if(-not($btsCUs -match "CU$($BizTalk_CU_VERSION)" )) {

        LogWarning "BizTalk CU$($BizTalk_CU_VERSION) is not installed in the current BizTalk group. Please install Microsoft BizTalk Server 2013 CU$($BizTalk_CU_VERSION)"        

    }

    LogInfo "The following Host Integration Server CUs are installed in the current BizTalk group."

    $hisCUs | % { LogInfo $_ }         

    if(-not($hisCUs -match "CU$($HostIntegrationServer_CU_VERSION)" )) {

        LogWarning "Host Integration Server CU$($HostIntegrationServer_CU_VERSION) is not installed in the current BizTalk group. Please install Host Integration Server CU$($HostIntegrationServer_CU_VERSION)"        

    }

 

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null 

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null  

    $regEntry = gi -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\BizTalk Server\3.0\Administration"

    $mgmtServer = "localhost"

    if($regEntry -ne $null) { $mgmtServer = $regEntry.GetValue("MgmtDBServer") }

    $ServerConnection =new-object "Microsoft.SqlServer.Management.Common.ServerConnection" $mgmtServer

    $msgBoxServer = $mgmtServer

    try {

        $msgBoxServer = $ServerConnection.ExecuteScalar("select [SubscriptionDBServerName] from [BizTalkMgmtDb].[dbo].[adm_Group]")

    }catch {}

    $ServerConnection.Disconnect()

    $ServerConnection =new-object "Microsoft.SqlServer.Management.Common.ServerConnection" $msgBoxServer

    $Server=New-Object "Microsoft.SqlServer.Management.Smo.Server" $ServerConnection

    if($Server.ProductLevel -ne "SP$($SQL_SP_VERSION)") {

        LogWarning "SQL Server SP$($SQL_SP_VERSION) is not installed. Please install SQL Server 2012 SP$($SQL_SP_VERSION)"

        LogInfo "The currently installed service pack version for SQL Server 2012 is $($Server.ProductLevel)."        

    }

    $ServerConnection.Disconnect()

    $Server = $null

    

    $dotnetVersion="$($PSVersionTable.CLRVersion.Major).$($PSVersionTable.CLRVersion.Minor)"

    LogInfo "The currently installed .NET Framework version is $dotnetVersion."        

    if($PSVersionTable.CLRVersion.Major -lt 4) {

        LogWarning ".NET Framework version is less than V4.x. Please install .NET Framework 4.x."        

    }

    LogInfo "Completed cumulative updates analysis..."

}

 

function ValidateDeploymentSources() {    

    CheckCUs

    LogInfo "Starting deployment pack validation..."

    $result = $true

    

    if($missingFile -ne $null) {

        LogInfo "The following file(s) must be present in $pwd. Please run deploy.ps1 using exportCSV flag on a computer that has MS Excel installed to generate the CSV files, place the CSVs into this directory & then re-run deploy.ps1 on this computer."

        LogInfo $missingFile        

        LogInfo "Terminating execution..."

        return  $false

    }

    

    

    

    Get-ChildItem -Path $msiSourcesRootFolder -Recurse | ?{ $_.PSIsContainer } | % {

        $packageName = $_.Name

        $deployable = $global:deploymentPlanMap | where { $_.Package -eq $packageName }

        $version = $null

        $msiName = $_.GetFiles("$packageName*.msi")

        if(-not([System.String]::IsNullOrWhiteSpace($msiName))) {

            $version = $msiName.Name.Substring($packageName.Length+1)

            if($version  -match "\d+[.]\d+[.]\d+") { $version = $version.Substring(0, $version.Length-4) }

        }

        $settingsFile = $_.GetFiles("*.xml")

        $deployBat = $_.GetFiles("$DEPLOY_BAT") 

        if($msiName -eq $null -or $msiName.Count -eq 0) {

            if($deployable -ne $null) {

                LogWarning "$packageName : The folder $($_.FullName) is missing an MSI. Please verify/rectify the release folder structure with all relevant drop files and re-run the script."            

            }

        }

        if($settingsFile -eq $null -or $settingsFile.Count -eq 0) {

            if($deployable -ne $null) {

                LogWarning "$packageName : The folder $($_.FullName) is missing a Settings XML file. Please verify/rectify the release folder structure with all relevant drop files and re-run the script."

            }

        }

        if($deployBat -eq $null -or $deployBat.Count -eq 0) {

            if($deployable -ne $null) {

                LogWarning "$packageName : The folder $($_.FullName) is missing '$DEPLOY_BAT'. Please verify/rectify the release folder structure with all relevant drop files and re-run the script."

            }

        }

        if(($msiName.Count -gt 1) -or ($settingsFile.Count -gt 1)) {

            $mesg = "$packageName : The folder $($_.FullName) contains more than one MSI/Settings files for $packageName. Deploy requires a single MSI & single Settings file within the subfolder having the BizTalk application name. Please  rectify the release folder structure with all relevant drop files and re-run the script."

            LogError $mesg

            $result = $false

        }

        if($version -ne $null) {

            if(-not($backoutMode)) {

                $hit = $global:packagesMap | where { $_.Package -eq $packageName -and $_.RolloutVersion -eq $version}

            } else {

                $hit = $global:packagesMap | where { $_.Package -eq $packageName -and $_.BackoutVersion -eq $version}

            }

            if($hit -eq $null) {

                LogWarning "$packageName : $msiName found in $($_.FullName) doesn't have a corresponding entry in deploy.xlsx '$WORKSHEET_PACKAGES' tab . Please rectify & reconcile $msiSourcesRootFolder and then re-run this script."

                $result = $false

            }

        }

    }

 

    

    $global:deploymentPlanMap | 

        where { @($_.PSObject.Properties)[2].Value -in ("Install","Deploy") } | select -ExpandProperty Package | 

        % {

            $msiName = $null

            $packageName = $_

            $pkgFolder = [System.IO.Path]::Combine($msiSourcesRootFolder, $packageName)

            if(-not(test-path $pkgFolder)) {

                LogError "$packageName : Could not find '$pkgFolder'. Installer files for the application $packageName must be contained in a subfolder with the application name. Please setup $msiSourcesRootFolder structure and pre-requisites correctly & then run this script again."                

                $result = $false

            }

            $hit = ($global:packagesMap | where { $_.Package -eq $packageName })[0]

            if($hit -ne $null) {

                if($backoutMode) { $msiVersion = $hit.BackoutVersion } else { $msiVersion = $hit.RolloutVersion }

                $msiName = "$($hit.Package)-$($msiVersion).msi"

                $msiFile = Test-Path -Path "$pkgFolder\*.msi"

                if(-not($msiFile)) {

                    LogError "$packageName : The folder $($pkgFolder) is missing $($msiName) which is required to install $packageName. Please rectify the release folder structure with all relevant drop files and re-run the script."

                    $result = $false

                }

            }

        }

    

    $global:deploymentPlanMap | 

        where { @($_.PSObject.Properties)[2].Value -in ("Deploy") } | select -ExpandProperty Package | 

        % {

            $packageName = $_

            $pkgFolder = [System.IO.Path]::Combine($msiSourcesRootFolder, $packageName)

            $settingsFile = Test-Path -Path "$pkgFolder\*.xml"

            if(-not($settingsFile)) {

                LogError "$packageName : The folder $($pkgFolder) is missing the target environment specific Settings XML configuration file for $packageName. Please rectify the release folder structure with all relevant drop files and re-run the script."

                $result = $false

            }

            $batFile = Test-Path -Path "$pkgFolder\$DEPLOY_BAT"

            if(-not($batFile)) {

                LogError "$packageName : The folder $($pkgFolder) is missing '$DEPLOY_BAT' file required to deploy $packageName into BizTalk management database. Please rectify the release folder structure with all relevant drop files and re-run the script."

                $result = $false

            }

        }

 

    $global:deploymentPlanMap | 

        where { @($_.PSObject.Properties)[2].Value -in ("ReDeploy") } | select -ExpandProperty Package | gu | 

        % {

            $packageName = $_

            $hit = $installedMap[$packageName]

 

            if($hit -eq $null) {

                LogWarning "$packageName : The BizTalk application $($packageName) is marked for 'ReDeploy' in deploy.xlsx. However the application is not installed, hence it cannot be re-deployed. Please include the MSI for this application in the release pack and update deploy.xlsx to 'Deploy' or 'Install' this application."

                $result = $false

            }            

        }

 

    LogInfo "Completed deployment pack validation...Validation Result = $result"

    return $result

}

 

 

function PostDeployBizTalkValidation() {

    $notDeployedApps = @()

    $catalog.Refresh()

    $global:deploymentPlanMap | 

        where { @($_.PSObject.Properties)[2].Value -in ("ReDeploy","Deploy") } | select -ExpandProperty Package | 

        % {

            $msiName = $null

            $packageName = $_

            try { 

                $deployedApp = $catalog.Applications | where {$_.Name -eq $packageName}               

                if(-not($deployedApp)) {

                    $notDeployedApps = $notDeployedApps + $packageName

                } else {

                    if($deployedApp.ReceivePorts.Count -eq 0 -and $deployedApp.Assemblies.Count -eq 0) {

                        $notDeployedApps = $notDeployedApps + $packageName

                    }

                }

            }catch{}

            

        }

    if($notDeployedApps.Count -gt 0) {

        LogWarning "*****The following BizTalk applications failed to deploy into management database. Please verify the corresponding *_Deploy.log files and deploy them manually."

        $notDeployedApps | % { LogWarning $_ }

    }

}

 

function DeleteFromGAC($application, $version) {    

    Write-Host -ForegroundColor Yellow "Server $($env:COMPUTERNAME): Uninstalling $version of $application from GAC"

    $DataStamp = get-date -Format yyyyMMddTHHmmss

    $gaclogFile = 'GAC-Remove-{0}-{1}.log' -f $application,$DataStamp

    $appRootPath = [System.IO.Path]::Combine($prgFilesLoc, $application, $PROJECT_VERSION)

    $gacUtilPath = [System.IO.Path]::Combine($appRootPath, "Deployment", "Framework", "DeployTools", "GacUtil.exe")

 

    $GACArguments = "/uf"

    $procHandle = Start-Process $gacUtilPath -ArgumentList $GACArguments -Wait -NoNewWindow -PassThru -RedirectStandardOutput $gaclogFile -RedirectStandardError $gaclogFile

    if($procHandle.HasExited) {

        if($procHandle.ExitCode -gt 0) { $colour = "Red" } else { $colour = "Green" }

        Write-Host -ForegroundColor $colour "Finished Un-GACing $application version $version on server $($env:COMPUTERNAME)" 

        if($procHandle.ExitCode -gt 0) { Write-host -ForegroundColor Red "Un-GACing failed on server $($env:COMPUTERNAME) with exit code $($procHandle.ExitCode). Please check 'My Documents' folder for the error log $gaclogFile to troubleshoot." }

    } else {

        Write-Host -ForegroundColor Red "Un-GACing of $application version $version was triggered on server $($env:COMPUTERNAME) but completion status could not be ascertained. Please manually delete it from GAC."

    }

}

 

 

function Launch([string]$exePath, [string]$exeArgs) {

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo

    $pinfo.FileName = $exePath

    $pinfo.RedirectStandardError = $true

    $pinfo.RedirectStandardOutput = $true

    $pinfo.UseShellExecute = $false

    $pinfo.Arguments = $exeArgs

    $p = New-Object System.Diagnostics.Process

    $p.StartInfo = $pinfo

    $p.Start()

    $p.WaitForExit()

    $stdout = $p.StandardOutput.ReadToEnd()

    $stderr = $p.StandardError.ReadToEnd()

    Write-Host "Output: $stdout"

    Write-Host $stderr

    Write-Host "exit code: " + $p.ExitCode

}

 

function LogInfo($mesg) {

    Write-Host $mesg

    $mesg | Out-File -FilePath $runLogFile -Append -Encoding utf8 

}

 

function LogWarning($mesg) {

    Write-Warning $mesg

    $mesg | Out-File -FilePath $runWarnLogFile -Append -Encoding utf8 

}

 

function LogError($mesg, $exception) {

    Write-Host -f Red "ERROR: $mesg"

    $mesg | Out-File -FilePath $runErrLogFile -Append -Encoding utf8 

    if($exception -ne $null) {

        Write-Host -f Red "Error details: $exception"

        "Error details: $exception" | Out-File -FilePath $runErrLogFile -Append -Encoding utf8 

    }    

}

 

function StopApplications() {

    $apps = $global:deploymentPlanMap | 

        where { @($_.PSObject.Properties)[2].Value -eq "Undeploy" } | select -ExpandProperty Package | gu 

    

    $apps    | % {

        $packageName = $_

        $app = $catalog.Applications | where { $_.Name -eq $packageName }

        

        if($app -ne $null)  {

            $applName = $app.Name

            try {                  

            

                if($app.Status -ne "Stopped" -and $app.Status -ne "NotApplicable" -and $app.Name.StartsWith("RSA.ESB", [System.StringComparison]::CurrentCultureIgnoreCase)) {

                    LogInfo "Stopping application $applName"  

                    $app.Stop([Microsoft.BizTalk.ExplorerOM.ApplicationStopOption]::StopAll -bor [Microsoft.BizTalk.ExplorerOM.ApplicationStopOption]::StopReferencedApplications)

                    $catalog.SaveChanges()

                    LogInfo "Stopped application $applName"

                }            

            }

            catch [System.Exception] {

                $handled = $false

                $schedulerDLLPath = "${env:ProgramFiles(x86)}\Microsoft BizTalk Server 2013\Microsoft.BizTalk.Scheduler.dll"

                if($_.Exception.Message.Contains("Could not load file") -and

                    $_.Exception.Message.Contains("Microsoft.BizTalk.Scheduler")) {

                    [System.Reflection.Assembly]::Load(“System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a”)

                    if(Test-Path "$schedulerDLLPath") {

                        $publish = New-Object System.EnterpriseServices.Internal.Publish

                        $publish.GacInstall(“$schedulerDLLPath”)

                        $handled = $true

                    }

                }

                if($_.CategoryInfo -and $_.CategoryInfo.Reason -eq "RuleEngineDeploymentNotDeployedException") {

                    $ruleSetName = $_.Exception.RuleSetName

                    LogWarning "Ruleset: $ruleSetName is not deployed. Please verify status and deploy it later."             

                }

                if(-not($handled) -and $btsAppStopErrCount -ge $btsAppStopErrThreshold) { 

                    $app = $catalog.Applications | where {$_.Name -eq $applName}

                    if($app -and $app.Status -eq "Stopped") {

                        LogInfo "Continuing with deployment after verifying that application $applName has stopped inspite of the error..."

                        $mesg = "There was an error while attempting to stop the application $applName. Degraded status to warning after verifying that the application has stopped. Error details - "

                        LogWarning $mesg

                        LogWarning $_.Exception.Message

 

                    } else {

                        LogError "Failed to stop the application $applName. If the application eventually failed to deploy please manually stop $applName choosing the option to terminate all instances and retry the deployment script."

                        #don't throw, ignore error

 

                    }    

                } 

                else {

                    if($btsAppStopErrCount -lt $btsAppStopErrThreshold) {                                        

                        $btsAppStopErrCount++

                        $catalog.Refresh()

                        LogInfo "Re-attempting to stop applications..."

                        StopApplications

                    }

                }

            }

        }

 

    }

    

}

 

 

function StartApplications() {

    $apps = $global:deploymentPlanMap | 

        where { @($_.PSObject.Properties)[2].Value -eq "Undeploy" } | select -ExpandProperty Package | gu 

    $startOrder = @()

    for($i=$apps.Count-1; $i -ge 0; $i--) { $startOrder = $startOrder + $apps[$i] }

    $startOrder    | % {

        $packageName = $_

        $app = $catalog.Applications | where { $_.Name -eq $packageName }

        

        if($app -ne $null) { 

            $applName = $app.Name

            try {                  

            

                if($app.Status -eq "Stopped" -and $app.Status -ne "NotApplicable" -and $app.Name.StartsWith("RSA.ESB", [System.StringComparison]::CurrentCultureIgnoreCase)) {

                    LogInfo "Starting application $applName"  

                    $app.Start([Microsoft.BizTalk.ExplorerOM.ApplicationStartOption]::StartAll)

                    $catalog.SaveChanges()

                    LogInfo "Started application $applName"

                }            

            }

            catch [System.Exception] {

                $handled = $false

                $schedulerDLLPath = "${env:ProgramFiles(x86)}\Microsoft BizTalk Server 2013\Microsoft.BizTalk.Scheduler.dll"

                if($_.Exception.Message.Contains("Could not load file") -and

                    $_.Exception.Message.Contains("Microsoft.BizTalk.Scheduler")) {

                    [System.Reflection.Assembly]::Load(“System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a”)

                    if(Test-Path "$schedulerDLLPath") {

                        $publish = New-Object System.EnterpriseServices.Internal.Publish

                        $publish.GacInstall(“$schedulerDLLPath”)

                        $handled = $true

                    }

                }

                if(-not($handled) -and $btsAppStartErrCount -ge $btsAppStartErrThreshold) {

                    LogError "Failed to start application: $applName. Please start it manually using BizTalk Administration Console. Details - $($_.Exception)"

                }

                else {

                    if($btsAppStartErrCount -lt $btsAppStartErrThreshold) {                                        

                        $btsAppStartErrCount++

                        $catalog.Refresh()

                        LogInfo "Re-attempting to start applications..."

                        StartApplications

                    }

                } 

            }

        }

 

    }

    

}

 

function StopHosts() {

    try {

        $enumOptions = New-Object System.Management.EnumerationOptions

        $enumOptions.ReturnImmediately = $false;

        $Scope = new-Object System.Management.ManagementScope "Root\MicrosoftBizTalkServer"

 

        $Query = New-Object System.Management.ObjectQuery "Select * from MSBTS_HostInstance where HostType=1"

        $searchObject = New-Object System.Management.ManagementObjectSearcher $Scope, $Query, $enumOptions

        $unclusteredHosts = @{}

        $clusteredHosts = @{}

        $clusterResource = @{}

        

        $searchObject.Get() | % { 

            $isDisabled = $_["IsDisabled"]

            $isClustered = ($_["ClusterInstanceType"] -ne 0)

            $serverName = $_["RunningServer"]

            $hostName = $_["HostName"]

            if(-not($isDisabled -or $isClustered)) {

                $unclusteredHosts.Set_Item($hostName , 1) 

            }

            if(-not($isDisabled) -and $isClustered) { $clusteredHosts.Set_Item($hostName, $serverName) }

        }

 

        $searchObject.Query = "Select * from MSBTS_Host where HostType=1"

        $searchObject.Get() | % { 

            $hostName = $_["Name"]

            if($clusteredHosts.ContainsKey($hostName)) { 

                $crgn = $_["ClusterResourceGroupName"]

                if(-not([System.String]::IsNullOrWhiteSpace($crgn))) {

                    $clusterResource.Set_Item($crgn, $clusteredHosts[$hostName])

                }

            } elseif($unclusteredHosts.ContainsKey($hostName)) {

                try{

                    LogInfo "Stopping host $($hostName)..."

                    $_.InvokeMethod("Stop",$null)

                    LogInfo "Successfully stopped host $($hostName)..."

                } catch [System.Exception] {

                    if($hostName -ne $null) { 

                        LogError "Failed to stop host $($hostName)" $_.Exception

                    } else { LogError "Failed to stop host. " $_.Exception }

                }

            }

           

        }

           

        $clusterResource.Keys | % { 

            $crgn = $_

            $remoteServer = $clusterResource[$crgn]

            if($remoteServer.Count -gt 1) { $remoteServer = $remoteServer[0] }

            $rSession = New-PSSession -ComputerName $remoteServer -EnableNetworkAccess

            $clusterFunctionDef = "function StopCluster { ${function:StopCluster} };"

 

            $clusterfuncCallBlock = {

                Param( $clusterFunctionDef, $crgn, $REMOTE_LOG_FOLDER, $DataStamp )

                . ([ScriptBlock]::Create($clusterFunctionDef))

                StopCluster $crgn $REMOTE_LOG_FOLDER $DataStamp

            }

 

            

            LogInfo "Stopping cluster role $($crgn)..." 

            $result = Invoke-Command -Session $rSession -ScriptBlock $clusterfuncCallBlock -ArgumentList $clusterFunctionDef, $crgn, $REMOTE_LOG_FOLDER, $DataStamp

            if($result) { "Successfully stopped cluster role $($crgn) ..." }

            else {

                $mesg = "Failed to stop cluster role $($crgn). For details please refer to the log file on $($remoteServer) of the pattern $($REMOTE_LOG_FOLDER)\$($DataStamp)\ClusterStop$($crgn)*.log."

                LogError $mesg

                LogInfo $mesg

            }

            Remove-PSSession $rSession

 

        }        

        LogInfo "All HostInstances stopped"   

    }

    catch [System.Exception] {

        LogError "Failure while stopping HostInstances. Please manually stop any hosts that are not started via BizTalk Administration Console. Details - $($_.Exception)"

    }  

}

 

function StopCluster($crgn, $REMOTE_LOG_FOLDER, $DataStamp) {

    try {

        $prefix = "$($REMOTE_LOG_FOLDER)\$($DataStamp)"

        if(-not(Test-Path -Path $prefix)) { ni -ItemType Directory -Force -Path $prefix | Out-Null }

        $logFile = "$prefix\ClusterStop{0}-{1}.log" -f $crgn,$DataStamp

        $errLogFile = "$prefix\ClusterStop{0}-{1}_ERRORs.log" -f $crgn,$DataStamp

        $cg = Get-ClusterGroup -Name "$($crgn)"

        if($cg -ne $null -and $cg.State -ne "Offline") { 

            $mesg = "Attempting to stop cluster role $($crgn)..."

            $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

            start-clustergroup -name "$($crgn)" -OutVariable outBuffer -ErrorVariable errBuffer 

            if([System.String]::IsNullOrEmpty($errBuffer)) { 

                $mesg = "Successfully stopped cluster role $($crgn) ..."

                Write-Host -f Green $mesg

                $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                $outBuffer | Out-File -FilePath $logFile -Append -Encoding utf8  

                return $true

            } else {

                $mesg = "ERROR*** Failed to stop the cluster role $($crgn). Please manually stop the role using Windows Failover Cluster Manager. Details - $($_.Exception)" 

                Write-host -ForegroundColor Red $mesg

                $mesg | Out-File -FilePath $errLogFile -Append -Encoding utf8  

                $errBuffer | Out-File -FilePath $errLogFile -Append -Encoding utf8  

                return $false

            }

        } elseif($cg -ne $null -and $cg.State -eq "Offline") { return $true }

    } catch [System.Exception] {                

            $mesg = "ERROR*** Failed to stop the cluster role $($crgn). Please manually stop the role using Windows Failover Cluster Manager. Details - $($_.Exception)" 

            Write-host -ForegroundColor Red $mesg

            $mesg | Out-File -FilePath $errLogFile -Append -Encoding utf8  

            return $false

    }

}

 

function StopHostInstances() {

    try {

        $enumOptions = New-Object System.Management.EnumerationOptions

        $enumOptions.ReturnImmediately = $false;

        $Scope = new-Object System.Management.ManagementScope "Root\MicrosoftBizTalkServer"

        $Query = New-Object System.Management.ObjectQuery "Select * from MSBTS_HostInstance where HostType=1"

        $searchObject = New-Object System.Management.ManagementObjectSearcher $Scope, $Query, $enumOptions

   

        $searchObject.Get() | % {   

            if($_["ServiceState"] -ne 1) {

                LogInfo "Stopping $($_['HostName']) on Server: $($_['RunningServer'])"

                $_.InvokeMethod("Stop",$null)

                LogInfo "Stopped $($_['HostName']) on Server: $($_['RunningServer'])"

            }    

        }

        LogInfo "All HostInstances stopped"   

    }

    catch [System.Exception] {

        LogError "Failure while stopping HostInstances. Please run the script again by specifying -stopOnly:`$true as a commandline parameter. Details - $($_.Exception)"

    }

  

}

 

function StartHostInstances() {

    try {

        $enumOptions = New-Object System.Management.EnumerationOptions

        $enumOptions.ReturnImmediately = $false;

        $Scope = new-Object System.Management.ManagementScope "Root\MicrosoftBizTalkServer"

        $Query = New-Object System.Management.ObjectQuery "Select * from MSBTS_HostInstance where HostType=1 and IsDisabled = FALSE and ServiceState=1"

        $searchObject = New-Object System.Management.ManagementObjectSearcher $Scope, $Query, $enumOptions

   

        $searchObject.Get() | % {   

            if($_["ServiceState"] -eq 1) { 

                LogInfo "Starting $($_['HostName']) on Server: $($_['RunningServer'])"

                $_.InvokeMethod("Start",$null)

                LogInfo "Started $($_['HostName']) on Server: $($_['RunningServer'])"

            }    

        }

        LogInfo "All HostInstances started"   

    }

    catch [System.Exception] {

        LogError "Failure while starting HostInstances. Please manually start any hosts that are not started via BizTalk Administration Console. Details - $($_.Exception)"

    }

  

}

 

function StartCluster($crgn, $REMOTE_LOG_FOLDER, $DataStamp) {

    try {

        $prefix = "$($REMOTE_LOG_FOLDER)\$($DataStamp)"

        if(-not(Test-Path -Path $prefix)) { ni -ItemType Directory -Force -Path $prefix | Out-Null }

        $logFile = "$prefix\ClusterStart{0}-{1}.log" -f $crgn,$DataStamp

        $errLogFile = "$prefix\ClusterStart{0}-{1}_ERRORs.log" -f $crgn,$DataStamp

        $cg = Get-ClusterGroup -Name "$($crgn)"

        if($cg -ne $null -and $cg.State -ne "Online") { 

            $mesg = "Attempting to start cluster role $($crgn)..."

            $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

            start-clustergroup -name "$($crgn)" -OutVariable outBuffer -ErrorVariable errBuffer 

            if([System.String]::IsNullOrEmpty($errBuffer)) { 

                $mesg = "Successfully started cluster role $($crgn) ..."

                Write-Host -f Green $mesg

                $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

                $outBuffer | Out-File -FilePath $logFile -Append -Encoding utf8  

                return $true

            } else {

                $mesg = "ERROR*** Failed to start the cluster role $($crgn). Please manually start the role using Windows Failover Cluster Manager. Details - $($_.Exception)" 

                Write-host -ForegroundColor Red $mesg

                $mesg | Out-File -FilePath $errLogFile -Append -Encoding utf8  

                $errBuffer | Out-File -FilePath $errLogFile -Append -Encoding utf8  

                return $false

            }

        } elseif($cg -ne $null -and $cg.State -eq "Online") { return $true }

    } catch [System.Exception] {                

            $mesg = "ERROR*** Failed to start the cluster role $($crgn). Please manually start the role using Windows Failover Cluster Manager. Details - $($_.Exception)" 

            Write-host -ForegroundColor Red $mesg

            $mesg | Out-File -FilePath $errLogFile -Append -Encoding utf8  

            return $false

    }

}

 

function StartHosts() {

    try {

        $enumOptions = New-Object System.Management.EnumerationOptions

        $enumOptions.ReturnImmediately = $false;

        $Scope = new-Object System.Management.ManagementScope "Root\MicrosoftBizTalkServer"

 

        #$Query = New-Object System.Management.ObjectQuery "Select * from MSBTS_HostInstance where HostType=1 and IsDisabled = FALSE and ClusterInstanceType=0"

        $Query = New-Object System.Management.ObjectQuery "Select * from MSBTS_HostInstance where HostType=1"

        $searchObject = New-Object System.Management.ManagementObjectSearcher $Scope, $Query, $enumOptions

        $unclusteredHosts = @{}

        $clusteredHosts = @{}

        $clusterResource = @{}

        #$searchObject.Get() | % { $unclusteredHosts.Set_Item($_["HostName"] , $_["RunningServer"]) }

 

        $searchObject.Get() | % { 

            $isDisabled = $_["IsDisabled"]

            $isClustered = ($_["ClusterInstanceType"] -ne 0)

            $serverName = $_["RunningServer"]

            $hostName = $_["HostName"]

            if(-not($isDisabled -or $isClustered)) {

                $unclusteredHosts.Set_Item($hostName , 1) 

            }

            if(-not($isDisabled) -and $isClustered) { $clusteredHosts.Set_Item($hostName, $serverName) }

        }

 

        $searchObject.Query = "Select * from MSBTS_Host where HostType=1"

        $searchObject.Get() | % { 

            $hostName = $_["Name"]

            if($clusteredHosts.ContainsKey($hostName)) { 

                $crgn = $_["ClusterResourceGroupName"]

                if(-not([System.String]::IsNullOrWhiteSpace($crgn))) {

                    $clusterResource.Set_Item($crgn, $clusteredHosts[$hostName])

                }

            } elseif($unclusteredHosts.ContainsKey($hostName)) {

                try{

                    LogInfo "Starting host $($hostName)..."

                    $_.InvokeMethod("Start",$null)

                    LogInfo "Successfully started host $($hostName)..."

                } catch [System.Exception] {

                    if($hostName -ne $null) { 

                        LogError "Failed to start host $($hostName)" $_.Exception

                    } else { LogError "Failed to start host. " $_.Exception }

                }

            }

           

        }

           

        $clusterResource.Keys | % { 

            $crgn = $_

            $remoteServer = $clusterResource[$crgn]

            if($remoteServer.Count -gt 1) { $remoteServer = $remoteServer[0] }

            $rSession = New-PSSession -ComputerName $remoteServer -EnableNetworkAccess

            $clusterFunctionDef = "function StartCluster { ${function:StartCluster} };"

 

            $clusterfuncCallBlock = {

                Param( $clusterFunctionDef, $crgn, $REMOTE_LOG_FOLDER, $DataStamp )

                . ([ScriptBlock]::Create($clusterFunctionDef))

                StartCluster $crgn $REMOTE_LOG_FOLDER $DataStamp

            }

 

            

            LogInfo "Starting cluster role $($crgn)..." 

            $result = Invoke-Command -Session $rSession -ScriptBlock $clusterfuncCallBlock -ArgumentList $clusterFunctionDef, $crgn, $REMOTE_LOG_FOLDER, $DataStamp

            if($result) { "Successfully started cluster role $($crgn) ..." }

            else {

                $mesg = "Failed to start cluster role $($crgn). For details please refer to the log file on $($remoteServer) of the pattern $($REMOTE_LOG_FOLDER)\$($DataStamp)\ClusterStart$($crgn)*.log."

                LogError $mesg

                LogInfo $mesg

            }

            Remove-PSSession $rSession

 

        }        

        LogInfo "All HostInstances started"   

    }

    catch [System.Exception] {

        LogError "Failure while starting HostInstances. Please manually start any hosts that are not started via BizTalk Administration Console. Details - $($_.Exception)"

    }

  

}

 

function IISReset($verb, $REMOTE_LOG_FOLDER, $DataStamp) {    

    $prefix = "$($REMOTE_LOG_FOLDER)\$($DataStamp)"

    if(-not(Test-Path -Path $prefix)) { ni -ItemType Directory -Force -Path $prefix | Out-Null }

    $logFile = "$prefix\IISReset{0}-{1}.log" -f $packageName,$DataStamp

    

    $mesg = "Server $($env:COMPUTERNAME): Executing IIS reset $verb"

    Write-Host $mesg

    $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

 

    $arguments = "/$verb"

        

    $procHandle = Start-Process "iisreset" -ArgumentList $arguments -Wait -NoNewWindow -PassThru -RedirectStandardOutput $logFile

    if($procHandle.HasExited) {

        if($procHandle.ExitCode -gt 0) { $colour = "Red" } else { $colour = "Green" }

        $mesg = "Finished iis $verb on server $($env:COMPUTERNAME)" 

        Write-Host -ForegroundColor $colour $mesg

        $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        if($procHandle.ExitCode -gt 0) { 

            $mesg = "iis $verb failed on server $($env:COMPUTERNAME) with exit code $($procHandle.ExitCode). Please check 'My Documents' folder for the error log $logFile to troubleshoot." 

            Write-host -ForegroundColor Red $mesg

            "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

        }

    } else {

        $mesg = "iis $verb was triggered on server $($env:COMPUTERNAME) but completion status could not be ascertained. Please verify un-installation manually via Control Panel. If the package is still present please retry the script."

        Write-Host -ForegroundColor Red $mesg

        "ERROR*** " + $mesg | Out-File -FilePath $logFile -Append -Encoding utf8  

    }

}

 

 

function DoParallelIISReset($verb) {

    $webServers = $global:serversMap | where { $_.Role -eq "Web"} | select -ExpandProperty Server

    LogInfo "Executing IIS $verb in parallel across servers $webServers"

    $allIISFunctionDefs = "function IISReset { ${function:IISReset} };"

    $IISfuncCallBlock = {

        Param( $allIISFunctionDefs, $verb, $REMOTE_LOG_FOLDER, $DataStamp )

        . ([ScriptBlock]::Create($allIISFunctionDefs))

        IISReset $verb $REMOTE_LOG_FOLDER $DataStamp

    }

    $webSession = New-PSSession -ComputerName $webServers -EnableNetworkAccess

 

    Invoke-Command -Session $webSession -ScriptBlock $IISfuncCallBlock -ArgumentList $allIISFunctionDefs,$verb, $REMOTE_LOG_FOLDER, $DataStamp 

    LogInfo "Finished executing IIS $verb in parallel across servers $webServers"

    Remove-PSSession $webSession

}

 

Function SSO-InstallApp{

param(

    [parameter(Mandatory=$true)]

    [ValidateNotNullOrEmpty()]    

    [string]$SSOApplicationName,

    [string]$ApplicationName

    

    )

    

    $SSOfilePath = "C:\Program Files (x86)\$ApplicationName\1.0"

    $logLocation = "$pwd\Logs\SSOInstallLog_$SSOApplicationName.txt"

    $SSODbConfigTool = $pwd.Path + "\BTSScnSSOApplicationConfig.exe"

    $SSOManageTool = "c:\Program Files\Common Files\Enterprise Single Sign-On\ssomanage.exe"

    

 

            $SSOXmlFile = $SSOfilePath + "\" + $SSOApplicationName + ".app.xml"

    $SSOKeyValueFile = $SSOfilePath + "\" + $SSOApplicationName + ".values.xml"

            

    $IsSSOInstallSuccess = $true

            $IsAnSSOAppCreated = $true

            $IsValuesAdded = $true

 

    Function SSO-PreprocXML{

    $xmlPreprocArgs = "/v /q /c /nologo /i `"$SSOfilePath\$SSOApplicationName.app.xml;$SSOFilePath\$SSOApplicationName.values.xml`" $PreProcSettings /e $Environment"

    $exe = (Start-Process -FilePath .\XmlPreprocess.exe -ArgumentList $xmlPreprocArgs -PassThru -Wait -NoNewWindow)

    if ($exe.exitcode -eq 0){"SSO XML Processing completed"}

    else{"Error in SSO xmlpreprocessing";#exit 6

}

    }

 

   # SSO-PreprocXML

    Write-Output "Trying to delete SSO App <$SSOApplicationName> from BizTalk SSO Store."

        try{$exe = (Start-Process -FilePath $SSOManageTool -ArgumentList "-deleteapp $SSOApplicationName" -PassThru -Wait -NoNewWindow)}

        catch{"Error: $_";#exit 5

}

        if (-Not $exe.exitcode -eq 0){"Error deleting SSO app";#exit 5

}

              

 

Write-Output "Now trying to create SSO App <$SSOApplicationName> in BizTalk SSO Store." 

        try{$exe = (Start-Process -FilePath $SSOManageTool -ArgumentList "-createapps `"$SSOXmlFile`"" -PassThru -Wait -NoNewWindow)}

        catch{"Error: $_";exit 5}

        if (-Not $exe.exitcode -eq 0){"Error: Failed to create the SSO App <$SSOApplicationName> in BizTalk SSO Store. ";exit 5}

 

 

#Looping the content from the SSO value XML file

                        

        [xml]$SSOKeyValueFileXML = Get-Content $SSOKeyValueFile

                        

                        Write-Output "Now trying to import the values from <$SSOKeyValueFile>...."

                        $LineNo = 0

                        foreach ($KeyValuePair in $SSOKeyValueFileXML.SSOKeyValues.keyvaluepair)

        {

                                    $LineNo = $LineNo + 1

            $Key = $KeyValuePair.key

            $Value = $KeyValuePair.value

            if($Value -eq $null){"Warning value is empty at $key"}

 

            try{

                                    if($value -match "\\"){$proc = Invoke-Command -scriptBlock {cmd /c $SSODbConfigTool -set $SSOApplicationName ConfigProperties `"$key`" $value}}

                                    else{Invoke-Command -scriptBlock {cmd /c $SSODbConfigTool -set $SSOApplicationName ConfigProperties `"$key`" `"$value`"}}

            if (-Not $proc){"<$LineNo> imported the value for $Key"}

            else{"<$LineNo> Error reported:

            $proc"#;exit 6

                        }

            }

            catch{"There was en error: $_ "#;exit 6

            }

            if (-Not $exe.exitcode -eq 0){"Error: Failed to add the key value pair $Key, $Value at key nr <$LineNo> in BizTalk SSO Store.";exit 6}

                                    

                        }

 

                                    Write-Output "Successfully added $LineNo values for the $SSOApplicationName SSO application."

}

 

Function SSO-BackupApp{

    Param ([string]$SSOApplicationName)

 

            #Initializing

            $Computer = gc env:computername

            Write-Output "Backup of SSO $SSOApplicationName at $Computer "

            $BackupSSO = $pwd.Path + "\Backup\SSOFiles"

    if (! (Test-Path $BackupSSO)) { mkdir $BackupSSO }

            $LogLocation = $pwd.Path + "\Logs\BackupSSO_" + $SSOApplicationName + ".txt"

    if (! (Test-Path $LogLocation)) {new-item $LogLocation -ItemType file -ErrorAction:SilentlyContinue}

            

            $SSOBackupUtility = $pwd.Path + "\BackupSSO.exe"

            $SSODbConfigTool = $pwd.Path + "\BTSScnSSOApplicationConfig.exe"

            $SSOVal = ""

            $SSOKeyListFileSuffix = "_CBDM"                 

            $SSOXMLFile = $BackupSSO + "\" + $SSOApplicationName + ".xml"

            $SSOTXTFile = $BackupSSO + "\" + $SSOApplicationName + ".txt"

            $SSOKeyListFile = $BackupSSO + "\" + $SSOApplicationName + $SSOKeyListFileSuffix + ".txt"

            $ReturnCode = 0

                        

            Write-Output "The Backup Folder -  $BackupSSO " 

            

            #Starting the processing

            

            #Taking backup of each application

            

                        Write-Output "Backup for SSO Application <$SSOApplicationName> in progress. Please wait..........." 

 

                        

                        if(  ([IO.File]::Exists($SSOXMLFile)) -and ([IO.File]::Exists($SSOTXTFile)) )

                        {

                                    Write-Output "export of application $SSOApplicationName already exists. Skipping the backup to preserve initial copy"

                                    Exit 0 

                        }

                        else

                        {

                                    Write-Output "Calling utility BackupSSO.exe" 

                                    Write-Output "=============================" 

                                    

                                    #Caling the BackupSSO.exe utility

                                    Write-Output "$SSOBackupUtility $SSOApplicationName $BackupSSO $LogLocation $SSOKeyListFileSuffix "

                                    & $SSOBackupUtility $SSOApplicationName $BackupSSO $LogLocation $SSOKeyListFileSuffix | out-null

                                    

                                               #Based on the SSO Key List file generated by the BackupSSO.exe get the values using BTSScnSSOApplicationConfig.exe

                                               #Loop through the key List file generated by BackupSSO.exe

                                               

                                                if([IO.File]::Exists($SSOKeyListFile) -and  ([IO.File]::Exists($SSOXMLFile)) )

                                               {

                                                           Write-Output "Found the SSO Key list file <$SSOKeyListFile>"

                                                           Write-Output "Found the SSO XML file <$SSOXMLFile>" 

                                                           Write-Output "Trying to backup the values based on SSO Key List file....." 

                                                           #if file already exist then delete the SSO Value file

                                                           if([IO.File]::Exists($SSOTXTFile))

                                                           {

                                                                       Write-Output "A previous SSO value file found and trying to remove it...." 

                                                                       Remove-Item $SSOTXTFile

                                                           }

                                                           foreach ($SSOKey in get-content $SSOKeyListFile)

                                                           {

                                                                       $SSOVal = & $SSODbConfigTool '-get' $SSOApplicationName "ConfigProperties" $SSOKey

                                                                       $ChkValue = "Property <" + $SSOKey + "> = <--NO_VALUE_FOUND_FOR_PROPERTY-->"

                                                                       if( $SSOVal -ne $ChkValue)

                                                                       {

                                                                                   $SSOVal = $SSOVal.Replace("Property <" + $SSOKey + "> = <", $SSOKey + ";")

                                                                                   $SSOVal = $SSOVal.SubString(0,$SSOVal.Length -1)

                                                                                   Write-Output $SSOVal >> $SSOTXTFile

                                                                       }

                                                                       else

                                                                       {

                                                                                   Write-Output "Warning : No value set for the SSO Key <$SSOKey>."

                                                                       }

                                                           }

                                                           Write-Output "Backup completed succcessfully!!" 

                                               

                                                           Write-Output "Now trying to remove the Key List file <$SSOKeyListFile>......."

                                                           Write-Output "" 

                                                           Remove-Item $SSOKeyListFile

                                               }

                                               else

                                               {

                                                           Write-Output "Could not find the SSO Key list file <$SSOKeyListFile>...Exiting"

                                                           $ReturnCode = 1

                                                           Write-Output "Script Exit Code : $ReturnCode"

                                                           Exit $ReturnCode                                            }

                                               

                                    }

 

 

            Write-Output "Post processing backup .txt file to set triple quotes needed on import" 

            

            $src = '"'

            $dst = '"""'

            $textfile = $BackupSSO + "\" + $SSOApplicationName + ".txt"

            $content = Get-Content -path $textfile

            

            

            Try

                        {                   

                        $content | foreach-object {$_ -replace $src, $dst} | Set-Content $textfile

                        }

            Catch 

                        {

                        Write-Output "Could not replace in $textfile.." 

                        }

            

 

Write-Output "Convert txt to xml and cleanup"

 

$outfile = "$BackupSSO\$SSOApplicationName.values.xml"

$SSOKeyValueFile = "$BackupSSO\$SSOApplicationName.txt"

$int = 0

try { new-item $outfile -ItemType file -ErrorAction:Stop }

catch { Write-Output "$outfile already exist - Exiting"

            $ReturnCode = 1

            Write-Output "Script Exit Code : $ReturnCode"

            Exit $ReturnCode }

 

try {

$XmlWriter = New-Object System.XML.XmlTextWriter($outfile,$Null)

 

$xmlWriter.Formatting = 'Indented'

 

$XmlWriter.WriteStartDocument()

$xmlWriter.WriteStartElement("SSOKeyValues")

 

 

                                    foreach ($SSOTxtFileLine in get-content $SSOKeyValueFile)

                        {

                                    $SSOKey = $SSOTxtFileLine.Split(";")[0].Trim()

                                    $SSOValue = $SSOTxtFileLine.Split(";")[1]

                                    $xmlWriter.WriteStartElement("keyvaluepair")

            $xmlWriter.WriteAttributeString("key","$SSOKey")

            $xmlWriter.WriteAttributeString("value","$SSOValue")

            $xmlWriter.WriteEndElement()

                            

                                    }

                                    

        

$xmlWriter.WriteEndElement

$xmlWriter.WriteEndDocument()

$xmlWriter.Finalize

$xmlWriter.Flush

$xmlWriter.Close()

}

catch { Write-Output "Error generating xml"

            $ReturnCode = 1 }

 

Remove-Item $SSOTXTFile -ErrorAction:SilentlyContinue

try { Rename-Item $SSOXMLFile $BackupSSO\$SSOApplicationName.app.xml -ErrorAction:Stop }

catch { Write-Output "$SSOXMLFile not generated or $SSOApplicationName.app.xml already exist"

            $ReturnCode = 1 }

 

            

 

 

#Logging the exit code and returning

 

Write-Output "Function Exit Code : $ReturnCode"

 

}

 

 

 

function ExecuteActionOnServers($packageName, $action, $servers) {

            $stepResult = $true

    switch ($action)

    {

        "Undeploy" { $stepResult = Undeploy $packageName }

        "Uninstall" { 

            $hit = ($global:packagesMap | where { $_.Package -eq $packageName })[0]

            $rolloutVersion = $hit.RolloutVersion

            $backoutVersion = $hit.BackoutVersion

            if($backoutMode) { $targetVersion = $rolloutVersion } else { $targetVersion = $backoutVersion }

#            $stepResult = DoParallelUninstall $packageName $targetVersion $servers

        }

        "Install" {

            $hit = ($global:packagesMap | where { $_.Package -eq $packageName })[0]

            $rolloutVersion = $hit.RolloutVersion

            $backoutVersion = $hit.BackoutVersion

            if($backoutMode) { $targetVersion = $backoutVersion } else { $targetVersion = $rolloutVersion }

#            $stepResult = DoParallelInstall $packageName $targetVersion $servers

        }

        "QuickDeploy" { $stepResult = DoQuickDeploy $packageName $servers }

#        "IISQuickDeploy" { $stepResult = DoIISQuickDeploy $packageName $servers }

 #       "WebDeploy" { $stepResult = DoWebDeploy $packageName $servers }

        "Deploy" { $stepResult = DeployTrue $packageName }

        "ReDeploy" { $stepResult = ReDeploy $packageName }        

    }

            return $stepResult

}

 

function ExecuteAction($packageName, $parallelExec,$release, $actions, $mappedServers) {

    $execAction = $null

    $execServers = @()

    if($release -eq "Y"){

    if($parallelExec -eq "No") {

        for($i=0; $i -lt $actions.count; $i++) { 

            if(-not([System.String]::IsNullOrEmpty($actions[$i]))) {

                $execAction = $actions[$i]

                $execServers = $mappedServers[$i]

                break

            }

        }

        return ExecuteActionOnServers $packageName $execAction $execServers

    }    

    else {

        $beg=0;

                        $parallelActionsFinalResult = $true 

        $actions | select -Unique | % {

                                    if(-not([System.String]::IsNullOrEmpty($_))) {

                

                                               $currAction = $_.Trim()

                $execServers = @()

                                               for($i=0; $i -lt $actions.count; $i++) { if($actions[$i].Trim() -eq $currAction) { $execServers += $mappedServers[$i] } }

                Write-Host -f Yellow $packageName, $release, $currAction

                if($release -ne 'Y'){

                $currAction = 'Ignore'

                }

               

                if($currAction -ne 'Web'){

               

                if($currAction -eq 'Undeploy'){

                 $tempResult = ExecuteActionOnServers $packageName $currAction $execServers

                $tempResult = ExecuteActionOnServers $packageName Uninstall $execServers

                

                }

                 if($currAction -eq 'Uninstall'){

                 $tempResult = ExecuteActionOnServers $packageName $currAction $execServers

               

                

                }

                if($currAction -eq 'Deploy'){

                $tempResult = ExecuteActionOnServers $packageName Install $execServers

                 $tempResult = ExecuteActionOnServers $packageName $currAction $execServers

                }

                 if($currAction -eq 'Quickdeploy'){

                $tempResult = ExecuteActionOnServers $packageName Uninstall $execServers

                $tempResult = ExecuteActionOnServers $packageName Install $execServers

                 $tempResult = ExecuteActionOnServers $packageName $currAction $execServers

                }

                 if($currAction -eq 'Webdeploy'){

                 $tempResult = ExecuteActionOnServers $packageName Uninstall $execServers

                $tempResult = ExecuteActionOnServers $packageName Install $execServers

                 $tempResult = ExecuteActionOnServers $packageName $currAction $execServers

                }

                }

                 if($currAction -eq 'Web'){

                 

                 }

                 

                                                $parallelActionsFinalResult = $parallelActionsFinalResult -and $tempResult  

             

                                    }

                        }

                        return $parallelActionsFinalResult

                                                   

    }}

}

 

function ExecuteDeploymentPlan() {

    $columnCount = $global:deploymentPlanMap[0].PSObject.Properties.Name.Count

    if($columnCount -eq 0) { $columnCount=12 }

    $mappedServers = $global:deploymentPlanMap[0].PSObject.Properties.Name[3..$columnCount]

 

    #copy MSIs

    $global:deploymentPlanMap.Package | sort | gu | % { CopyMSIsAcrossServers $_ $mappedServers}

 

    $global:deploymentPlanMap | % {

        $package = $_.Package 

        $parallelExec = $_.ParallelExecute   

        $release = $_.Release 

        $actions = $_.PSObject.Properties.Value[3..$columnCount]

        if(-not(ExecuteAction $package $parallelExec $release $actions $mappedServers)) { 

            Write-Host -f Yellow $package $parallelExec $release $actions

                                    #Write-Host -f Yellow "Terminating further execution due to the last error. Please rectify and then execute the deployment again. Please check the run log files - $runLogFile, $runErrLogFile & $runWarnLogFile to determine which applications were installed successfully. The deployment plan may potentially be reduced to remove those applications that were successfully installed to retry deployment only for the remainder applications..." 

                                    #break; 

                        }

    }

 

    #cleanup staging MSIs

    $global:deploymentPlanMap.Package | sort | gu | % { CleanupMSIsAcrossServers $_ $mappedServers}

}

 

 

 

$logPrefix = "$($LOCAL_LOG_FOLDER)\$($DataStamp)"

if(-not(Test-Path -Path $logPrefix)) { ni -ItemType Directory -Force -Path $logPrefix | Out-Null }

 

LogInfo "Started execution ID $DataStamp @ $(get-date) with state :"

LogInfo "Execution server: $btsMgmtDbServerName | BTS DB : $btsMgmtDbName | CSV flag: $exportCSV | Stop Only : $stopOnly | Start Only : $startOnly | Validate Only : $validateOnly | Skip Validation : $skipValidation | Keep Hosts & Apps Stopped : $keepHostsAndApplicationsStopped | Rollout MSI source : $msiSourcesRootFolder "

LogInfo "Execution log file : $runLogFile"

LogInfo "Execution errors log file : $runErrLogFile"

LogInfo "Execution warnings log file : $runWarnLogFile"

 

if($help) { 

    Usage

    Exit

}

 

[void] [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.ExplorerOM")

[void] [System.reflection.Assembly]::LoadWithPartialName("System.Management")

$regEntry = gi -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\BizTalk Server\3.0\Administration"    

if($regEntry -ne $null) { $btsMgmtDbServerName = $regEntry.GetValue("MgmtDBServer") }

$catalog = New-Object Microsoft.BizTalk.ExplorerOM.BtsCatalogExplorer

$catalog.ConnectionString = "SERVER=$btsMgmtDbServerName;DATABASE=$btsMgmtDbName;Integrated Security=SSPI"

 

$btsAppStopErrThreshold = 2

$btsAppStopErrCount = 0

$btsAppStartErrThreshold = 2

$btsAppStartErrCount = 0

 

 

if($exportCSV) {

    $configFilePath = Join-Path -Path $Pwd -ChildPath $configFileName

    if(-Not(Test-Path $configFilePath)) {

        Usage

        Exit

    }  

    $outDir = $Pwd

    $packagesCSV = "{0}_{1}.csv" -f (Get-Item -path $configFilePath).BaseName, $WORKSHEET_PACKAGES

    $deploymentPlanCSV = "{0}_{1}.csv" -f (Get-Item -path $configFilePath).BaseName, $WORKSHEET_Deployment_Plan

    $serversCSV = "{0}_{1}.csv" -f (Get-Item -path $configFilePath).BaseName, $WORKSHEET_Servers

 

    Write-Host -f White "Converting $configFilePath Excel to CSV."

    try {

    Excel2CSV $configFilePath $WORKSHEET_PACKAGES

    Write-Host -f White "Extracted $WORKSHEET_PACKAGES sheet to $packagesCSV."

    Excel2CSV $configFilePath $WORKSHEET_Deployment_Plan

    Write-Host -f White "Extracted $WORKSHEET_Deployment_Plan sheet to $deploymentPlanCSV." 

    Excel2CSV $configFilePath $WORKSHEET_Servers

    Write-Host -f White "Extracted $WORKSHEET_Servers sheet to $serversCSV."

    Write-Host -f Green "CSV files were extracted to the folder $Pwd. Please execute this script in production by placing the generated files in the same folder as the powershell script."

    } catch [System.Exception] {

        Write-Host -f Yellow "Failure while generating CSV files from $configFileName. Please check usage info below and try again. `nError Details:"

        Write-Host -f Red $_.Exception

        Usage

    }    

    Exit 

}

else {

    # CSVs already exported & available to use 

    $csvItems = Get-ChildItem -Filter "*.csv"

    if($csvItems -eq $null -or $csvItems.Count -eq 0) {

        Write-Host "Please ensure that the deploy_*.csv CSV files are present in the same folder as the powershell script. CSVs must be exported previously using the -exportCSV parameter. See usage."

        usage

        Exit

    }

    $packagesCsv = gi -Path "*$WORKSHEET_PACKAGES.csv"

    $deploymentPlanCsv = gi -Path "*$WORKSHEET_Deployment_Plan.csv"

    $serversCsv = gi -Path "*$WORKSHEET_Servers.csv"

        

    if($packagesCsv -eq $null -or -not($packagesCsv.Exists)) { $missingFile = "*$WORKSHEET_PACKAGES.csv " }

    if($deploymentPlanCsv -eq $null -or -not($deploymentPlanCsv.Exists)) { $missingFile += " *$WORKSHEET_Deployment_Plan.csv " }

    if($serversCsv -eq $null -or -not($serversCsv.Exists)) { $missingFile += " *$WORKSHEET_Servers.csv" }

    $global:packagesMap = Import-Csv -Path $packagesCsv 

    $global:deploymentPlanMap = Import-Csv -Path $deploymentPlanCsv

    $global:serversMap = Import-Csv -Path $serversCsv  

 

            gci -Path $msiSourcesRootFolder -Recurse | ?{ $_.PSIsContainer } | % {

        $packageName = $_.Name

        $settingsFile = $_.GetFiles("*.xml")

                        if($settingsFile -ne $null -and $settingsFile.Count -gt 0) {

            $settingsMap.Set_Item($packageName, $settingsFile.Name)

        }

            }

            

    if(-not(Get-Module -Name FailoverClusters)) {

        try{

            Install-WindowsFeature -name FailoverClusters -IncludeManagementTools -ErrorAction Ignore 

            Import-Module FailoverClusters -ErrorAction Ignore 

        } catch {

            LogWarning "In a multi server environment with Windows Failover Clusters & Roles, this script must be run on a server that has Windows Failover Cluster feature enabled. If no clusters are present in the BizTalk group this message can be ignored."

        }

    }

}

 

 

 

if($stopOnly) { 

#    DoParallelIISReset "stop"

    StopApplications

    StopHosts

    LogInfo "Terminating further script execution as -stopOnly parameter was specified."

    LogInfo "Finished execution ID $DataStamp @ $(get-date)"

    return 

}

if($startOnly) {

    StartApplications

    StartHosts

    #DoParallelIISReset "start"

    LogInfo "Terminating further execution since -startOnly parameter was specified."

    LogInfo "Finished execution ID $DataStamp @ $(get-date)"

    return

}

 

MapInstalledProducts

 

if($validateOnly) {

    ValidateDeploymentSources

    LogInfo "Terminating further execution since -validateOnly parameter was specified."

    LogInfo "Finished execution ID $DataStamp @ $(get-date)"

    return

}

 

#if(-not($skipValidation)) { ValidateDeploymentSources }

#DoParallelIISReset "stop"

StopApplications

StopHosts

 

ExecuteDeploymentPlan

PostDeployBizTalkValidation

 

# Finally start hosts 

if(-not($keepHostsAndApplicationsStopped)) {

    StartApplications

    StartHostInstances

   # DoParallelIISReset "start"

}

 

LogInfo "Finished execution ID $DataStamp @ $(get-date)"


 
