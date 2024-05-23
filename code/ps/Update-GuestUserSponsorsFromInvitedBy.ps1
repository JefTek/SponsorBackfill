<#
.Synopsis
   Update the Sponsors attribute to include the user who initially invited them to the tenant
.DESCRIPTION
   Update the Sponsors attribute to include the user who initially invited them to the tenant
.EXAMPLE
   Update-GuestUserSponsorsFromInvitedBy -All
.EXAMPLE
   Update-GuestUserSponsorsFromInvitedBy -UserId user1@contoso.com,user2@contoso.com
#>
function Update-GuestUserSponsorsFromInvitedBy {
    [CmdletBinding(SupportsShouldProcess,
        ConfirmImpact = 'High',
        DefaultParameterSetName = 'AllInvitedGuests')]
    param (

        # UserId of Guest User
        [Parameter(ParameterSetName = 'ByUsers')]
        [String[]]
        $UserId,
        # Enumerate and Update All Guest Users
        [Parameter(ParameterSetName = 'AllInvitedGuests')]
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
                $invitedBy = $null

                $splatArgumentsGetInvitedBy = @{

                    Method = 'Get'
                    Uri    = ("https://graph.microsoft.com/beta/users/" + $guestUser.Id + "/invitedBy")


                }

                $invitedBy = Invoke-MgGraphRequest @splatArgumentsGetInvitedBy

                Write-Verbose ("InvitedBy for Guest User {0}: {1}" -f $guestuser.DisplayName, $invitedBy.value.id)

                if ($null -ne $invitedBy) {
                    if ($guestUser.Sponsors -notcontains $invitedBy) {
                        write-verbose ("Sponsors does not contain the user who invited them!")

                        if ($PSCmdlet.ShouldProcess(("{0} - {1}" -f $guestUser.displayName, $guestUser.id), "Update Sponsors")) {
                            try {


                                $sponsorUrl = ("https://graph.microsoft.com/beta/users/{0}" -f $invitedBy.value.id)
                                $dirObj = @{"sponsors@odata.bind" = @($sponsorUrl) }
                                $sponsorsRequestBody = $dirObj | ConvertTo-Json


                                Update-MgBetaUser -UserId $guestUser.Id  -BodyParameter $sponsorsRequestBody


                            }
                            catch {

                                Write-Error $_

                            }
                        }


                    }
                }


            }
        }

    }

    end {

        Write-Verbose "Complete!"

    }
}