$lolbins = @("Atbroker.exe","Bash.exe","Bitsadmin.exe","Certutil.exe","Cmdkey.exe","Cmstp.exe","Control.exe","Csc.exe","Dfsvc.exe","Diskshadow.exe","Dnscmd.exe","Esentutl.exe","Eventvwr.exe","Expand.exe","Extexport.exe","Extrac32.exe","Findstr.exe","Forfiles.exe","Ftp.exe","Gpscript.exe","Hh.exe","Ie4uinit.exe","Ieexec.exe","Infdefaultinstall.exe","Installutil.exe","Makecab.exe","Mavinject.exe","Microsoft.Workflow.Compiler.exe","Mmc.exe","Msbuild.exe","Msconfig.exe","Msdt.exe","Mshta.exe","Msiexec.exe","Odbcconf.exe","Pcalua.exe","Pcwrun.exe","Presentationhost.exe","Print.exe","Reg.exe","Regasm.exe","Regedit.exe","Register-cimprovider.exe","Regsvcs.exe","Regsvr32.exe","Replace.exe","Rpcping.exe","Rundll32.exe","Runonce.exe","Runscripthelper.exe","Sc.exe","Schtasks.exe","Scriptrunner.exe","SyncAppvPublishingServer.exe","Verclsid.exe","Wab.exe","Wmic.exe","Wscript.exe","Xwizard.exe","Appvlp.exe","Bginfo.exe","Cdb.exe","csi.exe","dnx.exe","Dxcap.exe","Mftrace.exe","Msdeploy.exe","msxsl.exe","rcsi.exe","Sqldumper.exe","Sqlps.exe","SQLToolsPS.exe","te.exe","Tracker.exe","vsjitdebugger.exe","schtasks.exe")

$startday = (Get-Date) - (New-TimeSpan -Day 4)
Foreach($lolbin in $lolbins)
{
    Get-WinEvent -FilterHashtable @{logname="Microsoft-Windows-Sysmon/Operational";id=1;} | ?{ $_.message -match "`r`nImage: .*$lolbin`r`n" } | %{
#    Get-WinEvent -FilterHashtable @{logname="Microsoft-Windows-Sysmon/Operational";id=1;StartTime=$startday} | ?{ $_.message -match "`r`nImage: .*$lolbin`r`n" } | %{
        
        [regex]$regex = "(?i)`r`n(?<image>Image: .*$lolbin?)`r`n"
#        [regex]$regex = "(?i)`r`n(?<image>Image: .*$lolbin?)`r`n .*(?<args>CommandLine: .*?)`r`n"
#        $match = $regex.Match($_.message)
        $match1 = $regex.Match($_.message)

#        $Out = New-Object PSObject
#        $Out | Add-Member Noteproperty 'Binary' $lolbin
#        $Out | Add-Member Noteproperty 'Image' $match1.Groups["image"].value
#        $Out | Add-Member Noteproperty 'Args' $match.Groups["args"].value
#        $Out | fl

        [regex]$regex = "(?i)`r`n(?<args>CommandLine: .*?)`r`n"
        $match2 = $regex.Match($_.message)

        $Out = New-Object PSObject
        $Out | Add-Member Noteproperty 'Binary' $lolbin
        $Out | Add-Member Noteproperty 'Image' $match1.Groups["image"].value
        $Out | Add-Member Noteproperty 'Args' $match2.Groups["args"].value
        $Out | fl

    }
}