$FTE = Get-DynamicDistributionGroup "All Staff"

$rec = Get-Recipient -RecipientPreviewFilter $FTE.RecipientFilter -OrganizationalUnit $FTE.RecipientContainer

$rec.Count
