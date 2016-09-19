$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

#region Load Public Functions
Try {
    Get-ChildItem "$ScriptPath\Functions\Public" -Filter *.ps1 | Select -Expand FullName | ForEach {
        $Function = Split-Path $_ -Leaf
        . $_
    }
} Catch {
    Write-Warning ("{0}: {1}" -f $Function,$_.Exception.Message)
    Continue
}
#endregion Load Public Functions

#region Private Functions

#endregion Private Functions

#region Aliases
New-Alias -Name oac -Value Out-AreaChart
New-Alias -Name obac -Value Out-BarChart
New-Alias -Name obuc -Value Out-BubbleChart
New-Alias -Name occ -Value Out-ColumnChart
New-Alias -Name odc -Value Out-DoughnutChart
New-Alias -Name ofc -Value Out-FunnelChart
New-Alias -Name okc -Value Out-KagiChart
New-Alias -Name olc -Value Out-LineChart
New-Alias -Name opic -Value Out-PieChart
New-Alias -Name opoc -Value Out-PointChart
New-Alias -Name opyc -Value Out-PyramidChart
New-Alias -Name ospc -Value Out-SplineChart
New-Alias -Name osac -Value Out-StackedAreaChart
New-Alias -Name ostb -Value Out-StackedBarChart
New-Alias -Name oscc -Value Out-StackedColumnChart
#endregion Aliases

#region Load Type and Format Files

#endregion Load Type and Format Files

Export-ModuleMember -Alias * -Function *