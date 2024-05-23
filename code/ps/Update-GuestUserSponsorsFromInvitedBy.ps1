function Update-GuestUserSponsorsFromInvitedBy {
    [CmdletBinding(SupportsShouldProcess = $True)]
    param (

        # UserId of Guest User
        [String[]]
        $UserId,
        # Enumerate and Update All Guest Users
        [switch]
        $All
    )

    begin {
        $guestFilter = "(userType eq 'Guest' and ExternalUserState in ('PendingAcceptance', 'Accepted'))"

    }

    process {

        if ($All) {

            $GuestUsers = get-mguser -filter $guestFilter -ExpandProperty Sponsors
        }
        else {
            foreach ($user in $userId) {

                $GuestUsers += get-mguser -UserId $user -ExpandProperty Sponsors
            }
        }

        if ($null -eq $GuestUsers) {
            Write-Information "No Guest Users to Process!"
        }
        else {
            foreach ($guestUser in $GuestUsers) {
                #eval and update guest
            }
        }

    }

    end {

    }
}