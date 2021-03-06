Function Remove-OutlookSecurityPromptHC {
    [CmdLetBinding()]
    Param()

    if (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook') {
        Write-Verbose 'Found MS Outlook 2010'

        if (-not (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook\Security')) {
            New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook\Security' | Out-Null
        }
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook\Security' -Name ObjectModelGuard -Value 2
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook\Security' -Name PromptOOMSend -Value 2
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook\Security' -Name AdminSecurityMode -Value 3
        Write-Verbose 'Outlook warning suppressed'
    }

    if (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Office\12.0\Outlook') {
        Write-Verbose 'Found MS Outlook 2007'

        if (-not (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Office\12.0\Outlook\Security')) {
            New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Office\12.0\Outlook\Security' | Out-Null
        }
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\12.0\Outlook\Security' -Name ObjectModelGuard -Value 2
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\12.0\Outlook\Security' -Name PromptOOMSend -Value 2
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\12.0\Outlook\Security' -Name AdminSecurityMode -Value 3
        Write-Verbose 'Outlook warning suppressed'
    }
}

Remove-OutlookSecurityPromptHC -Verbose