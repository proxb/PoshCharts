Function Out-StackedColumnChart {
    <#
        .SYNOPSIS

        .DESCRIPTION

        .PARAMETER InputObject

        .PARAMETER XField

        .PARAMETER YField

        .PARAMETER Title

        .PARAMETER IncludeLegend

        .NOTES
            Name: Out-StackedColumnChart
            Author: Boe Prox
            Version History;
                1.0 //Boe Prox - 08/20/2016
                    - Initial Version

        .EXAMPLE
            Get-Process | Sort-Object WS -Descending | 
            Select-Object -First 10 | Out-StackedColumnChart -XField Name `
            -YField WS -Title 'Top 10 Processes By Working Set Memory'

        .EXAMPLE
            $Process1 = Get-WMIObject -Class Win32_PerfFormattedData_PerfProc_Process -Filter "Name != '_Total' AND Name != 'Idle' AND PercentProcessorTime != '0'"|
            Sort PercentProcessorTime -Descending  | Select -First 10
            Start-Sleep -Seconds 2
            $Process2 = Get-WMIObject -Class Win32_PerfFormattedData_PerfProc_Process -Filter "Name != '_Total' AND Name != 'Idle'" | Where {
                $_.Name -in $Process1.Name
            }
            @($Process1,$Process2) | Out-StackedColumnChart -XField Name `
            -YField PercentProcessorTime -Title 'Top 10 Processes By CPU Usage (if applicable)'

        .EXAMPLE
            Get-Process | Sort-Object WS -Descending | 
            Select-Object -First 10 | Out-StackedColumnChart -XField Name `
            -YField WS -ToFile "C:\users\proxb\desktop\File.jpeg" -IncludeLegend `
            -Title 'Top 10 Processes By Working Set Memory'

        .EXAMPLE
            Get-WMIObject -Class Win32_PerfFormattedData_PerfProc_Process -Filter "Name != '_Total' AND Name != 'Idle' AND PercentProcessorTime != '0'"|
            Sort PercentProcessorTime -Descending  | Select -First 10 | 
            Out-StackedColumnChart -XField Name -YField PercentProcessorTime -Title 'Top 10 Processes By CPU Usage (if applicable)' -IncludeLegend
    #>
    [cmdletbinding(
        DefaultParameterSetName = 'UI'
    )]
    Param (
        [parameter(ValueFromPipeline=$True)]
        $InputObject,
        [parameter()]
        [string]$XField,
        [parameter()]
        [string]$YField,
        [parameter()]
        [string]$Title = 'Test Title',
        [parameter()]
        [switch]$IncludeLegend,
        [parameter()]
        [string[]]$LegendText,
        [parameter()]
        [switch]$Enable3D,
        [parameter(ParameterSetName='File')]
        [ValidateScript({
            $UsedExt = $_ -replace '.*\.(.*)','$1'
            $Extensions = "Jpeg", "Png", "Bmp", "Tiff", "Gif", "Emf", "EmfDual", "EmfPlus"
            If ($Extensions -contains $UsedExt) {
                $True
            } 
            Else {
                Throw "The extension '$UsedExt' is not valid! Valid extensions are $($Extensions -join ', ')."
            }
        })]
        [string]$ToFile
    )
    Begin {
        #region Helper Functions
        function ConvertTo-Hashtable
        { 
            param([string]$key, $value) 

            Begin 
            { 
                $hash = @{} 
            } 
            Process 
            { 
                $thisKey = $_.$Key
                $hash.$thisKey = $_.$Value 
            } 
            End 
            { 
                Write-Output $hash 
            }

        }
        Function Invoke-SaveDialog {
            $FileTypes = [enum]::GetNames('System.Windows.Forms.DataVisualization.Charting.ChartImageFormat')| ForEach {
                $_.Insert(0,'*.')
            }
            $SaveFileDlg = New-Object System.Windows.Forms.SaveFileDialog
            $SaveFileDlg.DefaultExt='PNG'
            $SaveFileDlg.Filter="Image Files ($($FileTypes))|$($FileTypes)|All Files (*.*)|*.*"
            $return = $SaveFileDlg.ShowDialog()
            If ($Return -eq 'OK') {
                [pscustomobject]@{
                    FileName = $SaveFileDlg.FileName
                    Extension = $SaveFileDlg.FileName -replace '.*\.(.*)','$1'
                }
        
            }
        }
        #endregion Helper Functions
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization
        $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea         
        $ChartTypes = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
        $Chart.ChartAreas.Add($ChartArea)
        If ($PSBoundParameters.ContainsKey('Enable3D')) {
            $ChartArea.Area3DStyle.Enable3D=$True
            $ChartArea.Area3DStyle.Inclination = 50
        }

        $IsPipeline=$True
        $Data = New-Object System.Collections.ArrayList
        If ($PSBoundParameters.ContainsKey('InputObject')) {
            $Data.AddRange($InputObject)
            $IsPipeline=$False
        }
    }
    Process {
        If ($IsPipeline) {
            [void]$Data.Add($_)
        }
    }
    End {
        If ($Data[0] -is [array]) {
            $i=1
            ForEach ($Item in $Data) {
                $HashTable = $Item | ConvertTo-Hashtable -key $XField -value $YField
                $SeriesLabel = "Series$($i)"
                $Series = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Series -ArgumentList $SeriesLabel
                $Series.ChartType = $ChartTypes::StackedColumn
                $Chart.Series.Add($Series)            
                #region MSChart Build
                $Chart.Series[$SeriesLabel].Points.DataBindXY($HashTable.Keys, $HashTable.Values)
                $Chart.Series[$SeriesLabel]['ColumnLabelStyle'] = 'Outside'
                $Chart.Series[$SeriesLabel]['PixelPointWidth'] = 25
                If ($PSBoundParameters.ContainsKey('IncludeLegend')) {
                    $chart.Series[$SeriesLabel].LegendText = $LegendText[($i-1)]
                }
                #endregion MSChart Build
                $i++
            }
        }
        Else {
            $Series = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Series -ArgumentList 'Series1'
            $Series.ChartType = $ChartTypes::StackedColumn
            $Chart.Series.Add($Series) 
            $HashTable = $Data | ConvertTo-Hashtable -key $XField -value $YField
            #region MSChart Build
            $Chart.Series['Series1'].Points.DataBindXY($HashTable.Keys, $HashTable.Values)
            $Chart.Series['Series1']['ColumnLabelStyle'] = 'Outside'
            $Chart.Series['Series1']['PixelPointWidth'] = 25
                If ($PSBoundParameters.ContainsKey('IncludeLegend')) {
                    $chart.Series[$SeriesLabel].LegendText = $LegendText[($i-1)]
                }
            #endregion MSChart Build            
        }

        #region MSChart Configuration
        $Chart.Width = 700 
        $Chart.Height = 400 
        $Chart.Left = 10 
        $Chart.Top = 10
        $Chart.BackColor = [System.Drawing.Color]::White
        $Chart.BorderColor = 'Black'
        $Chart.BorderDashStyle = 'Solid'

        $ChartTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
        $ChartTitle.Text = $Title
        $Font = New-Object System.Drawing.Font @('Microsoft Sans Serif','12', [System.Drawing.FontStyle]::Bold)
        $ChartTitle.Font =$Font
        $Chart.Titles.Add($ChartTitle)    
        #endregion MSChart Configuration

        If ($PSBoundParameters.ContainsKey('IncludeLegend')) {
            #region Create Legend
            $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
            $Legend.IsEquallySpacedItems = $True
            $Legend.BorderColor = 'Black'
            $Chart.Legends.Add($Legend)
            #endregion Create Legend
        }
        $Chart.ChartAreas.axisx.LabelStyle.Angle = -45
        If ($PSBoundParameters.ContainsKey('ToFile')) {
            $Extension = $ToFile -replace '.*\.(.*)','$1'
            $Chart.SaveImage($ToFile, $Extension)
        } 
        Else {
            #region Windows Form to Display Chart
            $AnchorAll = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor 
                [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
            $Form = New-Object Windows.Forms.Form  
            $Form.Width = 740 
            $Form.Height = 490 
            $Form.controls.add($Chart) 
            $Chart.Anchor = $AnchorAll

            # add a save button 
            $SaveButton = New-Object Windows.Forms.Button 
            $SaveButton.Text = "Save" 
            $SaveButton.Top = 420 
            $SaveButton.Left = 600 
            $SaveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
            # [enum]::GetNames('System.Windows.Forms.DataVisualization.Charting.ChartImageFormat') 
            $SaveButton.add_click({
                $Result = Invoke-SaveDialog
                If ($Result) {
                    $Chart.SaveImage($Result.FileName, $Result.Extension)
                }
            }) 

            $Form.controls.add($SaveButton)
            $Form.Add_Shown({$Form.Activate()}) 
            [void]$Form.ShowDialog()
            #endregion Windows Form to Display Chart  
        }
    } 
}