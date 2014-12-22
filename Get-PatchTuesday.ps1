function Get-PatchTuesday {
<# 
 .Synopsis
  Retrives when the next Patch Tuesday is.

 .Description
  Retrives when the next Patch Tuesday is.

 .Parameter Month
  Numerical value of the month to retrive patch tuesday for.

 .Parameter Year
  Year to retrive patch tuesday for.
 
 .Example
  #Gets the current patch tuesday.
  Get-PatchTuesday

 .Example
  #Gets when Patch tuesday would have been in 1872 in January.
  Get-PatchTuesday -Month 1 -Year 1872

 .Inputs
  Does not accept pipeline input.

 .Outputs
  Date object of next patch tuesday.

 .Notes
  Thanks to Tim Curwick for the formula this was originally based off in his blog post here: http://www.madwithpowershell.com/2014/10/calculating-patch-tuesday-with.html

 .LINK
  about_WindowsUpdateModule

#>
    param($Month=((Get-Date).Month), $Year=((Get-Date).Year))

    $Date = Get-Date -Month $Month -Year $Year -Day 12
    $Date.AddDays(2 - [int]$date.DayOfWeek)
}

Export-ModuleMember Get-PatchTuesday