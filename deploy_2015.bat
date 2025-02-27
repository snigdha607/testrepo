@REM --------------------------------------------------------------------------------------------------------------
@REM Version 2.1, Release Date - 31/05/2016
@REM This BATCH file must be contained within a FOLDER with the same name as the BizTalk application being deployed
@REM Keep this bat file at same location where you keep your msi file
@REM --------------------------------------------------------------------------------------------------------------
@REM USAGE -
@REM *******
@REM To install the application and deploy into BizTalk management database the command line parameter should be TRUE as below.
@REM Deploy true
@REM To install the application ONLY and NOT deploy into BizTalk management database the command line parameter should be FALSE as below.
@REM Deploy false
@REM Usage -
@REM Deploy true => REMOVE BAM views, deploy the BizTalk app to MgmtDb, update BAM activities with latest definitions, deploy SSO settings and deploy SQL scripts (error and BAM configuration).
@REM Deploy false => Don’t deploy BAM, Don't deploy SSO settings, install into Programs & Features, deploy into GAC (not into MgmtDb), DEPLOY SQL scripts.
@REM Switches 
@REM S => Silent mode. Default: Silent is off. When S is specified it will not prompt for questions/answers
@REM P => Pause mode. Default: Pause if off. When P is specified it will pause at the end of execution.
@REM SKIPSQL => Don't execute SQL scripts. Default: SKIPSQL is off. When SKIPSQL is specified Sql scripts associated with autodeploy.proj are not executed at the end of deployment.
@ECHO off
@REM #######################################
@REM Defaults to prevent restarts on deploy
SET SKIPUNDEPLOY=true
SET SKIPHOSTINSTANCESRESTART=true
SET SKIPIISRESET=true
SET DEPLOYPDBSTOGAC=false
SET STARTAPPLICATIONONDEPLOY = false
SET STARTREFERENCEDAPPLICATIONSONDEPLOY = false
SET ENABLEALLRECEIVELOCATIONSONDEPLOY = false
SET AUTOTERMINATEINSTANCES = false
SET UNDEPLOYIISARTIFACTS = true
@REM #######################################
@REM #######################################
@REM Default for deployment file name extension for VS 2015
SET VS_EXT=_2015
@REM #######################################
SET MSI_DIR=%CD%
Echo batch1 folder is: %~dp0\config
Echo ======================
SET SETTINGS_DIR=D:\support\config
SET SETTINGS=settings
@ECHO:
@ECHO --- Settings folder-----
FOR %%I IN (.) DO SET APPLICATION=%%~nI%%~xI
FOR %%i in (*.msi) do SET MSI=%%~ni
@REM FOR %%i in (*.xml) do SET SETTINGS=%%~ni
SET TARGET="%PROGRAMFILES(x86)%\%APPLICATION%\1.0"
@ECHO:
@ECHO ----- STARTING %APPLICATION% deployment -----
@REM Initialise
SET SILENTMODE=0
SET PAUSEMODE=0
SET BTSUTILITY=1
SET RM_VIEWS=1
SET DEPLOY_MODE=MANUAL
IF /i {%2%}=={P} (
SET PAUSEMODE=1
)
IF /i {%2%}=={SKIPSQL} (
SET BTSUTILITY=0
)
IF /i {%2%}=={S} (
SET SILENTMODE=1
)
IF /i {%3%}=={P} (
SET PAUSEMODE=1
)
IF /i {%3%}=={SKIPSQL} (
SET BTSUTILITY=0
)
IF /i {%3%}=={S} (
SET SILENTMODE=1
)
IF /i {%4%}=={P} (
SET PAUSEMODE=1
)
IF /i {%4%}=={SKIPSQL} (
SET BTSUTILITY=0
)
IF /i {%4%}=={S} (
SET SILENTMODE=1
)
IF /i {%SILENTMODE%}=={1} (
goto :silent
)
@ECHO:
@ECHO 1. Please ensure that you have UNINSTALLED an earlier version of %APPLICATION% from Programs and Features before you continue.
@ECHO 2. Please ensure that %SETTINGS%.xml has been updated to reflect the current target deployment environment.
SET /P ANSWER=Have you verified and want to proceed? (Y/N)?
IF /i {%ANSWER%}=={y} (goto :yes)
IF /i {%ANSWER%}=={yes} (goto :yes)
goto :no
:silent
:yes
@ECHO:
@ECHO Step1: Installing %APPLICATION% BizTalk application using %MSI%.msi
msiexec /i %MSI%.msi /qb
SET BTDFMSBuildPath="%windir%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"
SET BT_DEPLOY_MGMT_DB=%1
SET PARMS=DeployBizTalkMgmtDB=%BT_DEPLOY_MGMT_DB%;Configuration=Server;SkipUndeploy=%SKIPUNDEPLOY%;SkipHostInstancesRestart=%SKIPHOSTINSTANCESRESTART%;SkipIISReset=%SKIPIISRESET%;DeployPDBsToGac=%DEPLOYPDBSTOGAC%;StartApplicationOnDeploy=%STARTAPPLICATIONONDEPLOY%;StartReferencedApplicationsOnDeploy=%STARTREFERENCEDAPPLICATIONSONDEPLOY%;EnableAllReceiveLocationsOnDeloy=%ENABLEALLRECEIVELOCATIONSONDEPLOY%;UndeployIISArchifacts=%UNDEPLOYIISARTIFACTS%
SET FilePath=%TARGET%
SET FilePath=%FilePath:"=%
IF EXIST "%FilePath%"\Deployment\AutoDeploy.proj (
SET DEPLOY_MODE=AUTO
)
IF EXIST "%FilePath%"\AutoDeploy.proj (
SET DEPLOY_MODE=AUTO
)
IF /i {%DEPLOY_MODE%}=={MANUAL} (
SET RM_VIEWS=0
SET BTSUTILITY=0
) 
@REM TODO: Remove temporaty Logging during RTM release
@echo DUMPING flags for debug...
@echo SILENTMODE=%SILENTMODE%
@echo PAUSEMODE=%PAUSEMODE%
@echo BTSUTILITY=%BTSUTILITY%
@echo RM_VIEWS=%RM_VIEWS%
@echo BT_DEPLOY_MGMT_DB=%BT_DEPLOY_MGMT_DB%
@echo DEPLOY_MODE=%DEPLOY_MODE%
@REM EndToDO
IF /i {%BT_DEPLOY_MGMT_DB%}=={false} (
@ECHO:
@ECHO Skipping Step2 to remove BAM views 
goto :step3
)
IF /i {%RM_VIEWS%}=={0} (goto :step3)
CD "%FilePath%"
IF /i "%CD%" NEQ "%FilePath%" (
%SYSTEMDRIVE%
)
CD "%MSI_DIR%"
:step3
@ECHO:
@ECHO Step3: Deploying application %APPLICATION% using settings %SETTINGS%.xml
%BTDFMSBuildPath% /t:Deploy "%FilePath%\Deployment\Deployment.btdfproj"  /p:%PARMS% /l:FileLogger,Microsoft.Build.Engine;logfile="%FilePath%\Deployment\DeployResults.txt" /p:ENV_SETTINGS="%SETTINGS_DIR%\%SETTINGS%.xml"
@ECHO:
@ECHO ----- FINISHED %APPLICATION% deployment -----
IF /i {%BT_DEPLOY_MGMT_DB%}=={false} (
@ECHO NOTE: In case of deploy false, SSO and BAM are not deployed.  
)
GOTO END
:no
@ECHO:
@ECHO ----- *ABORTING %APPLICATION% deployment* -----
:END
IF {%PAUSEMODE%}=={1} (
@PAUSE
)
