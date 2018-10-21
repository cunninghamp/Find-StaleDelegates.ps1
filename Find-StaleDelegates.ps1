<#
.SYNOPSIS
Find-StaleDelegates.ps1

.DESCRIPTION 
Sample script for locating stale delegate entries on Exchange Server mailboxes.

.EXAMPLE
.\Find-StaleDelegates.ps1.ps1

.LINK
https://exchangeserverpro.com

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

Change Log:
V1.00, 21/01/2015 - Initial version
#>

#Requires -Version 2

if (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -Registered -ErrorAction SilentlyContinue) {$mailboxes = @(Get-MailboxCalendarSettings | Where {$_.ResourceDelegates -ne $null})}
if (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -Registered -ErrorAction SilentlyContinue) {$mailboxes = @(Get-Mailbox | Get-CalendarProcessing | Where {$_.ResourceDelegates -ne $null})}

if ($mailboxes.count -eq 0)
{
    Write-Host "No mailboxes with resource delegates found."
    EXIT
}

foreach ($mailbox in $mailboxes)
{
    

	Write-Host -ForegroundColor White "*** Checking $($mailbox.Identity.Name)"

    $delegates = @($mailbox | Select -ExpandProperty:ResourceDelegates)

    foreach ($delegate in $delegates)
    {
        try
        {
            $getdelegate = Get-Recipient $delegate -ErrorAction STOP
            Write-Host -ForegroundColor Green "$($delegate.Name) is okay"
        }
        catch
        {
            Write-Host -ForegroundColor Yellow "$($delegate.Name) is stale"

        }
    }

}

Write-Host -ForegroundColor White "*** Finished"
