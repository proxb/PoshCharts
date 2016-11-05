Function Out-LineChart {
    <#
        .SYNOPSIS

        .DESCRIPTION

        .PARAMETER InputObject

        .PARAMETER XField

        .PARAMETER YField

        .PARAMETER YField1

        .PARAMETER YField2

        .PARAMETER Title

        .PARAMETER IncludeLegend

        .NOTES
            Name: Out-PieChart
            Author: Boe Prox
            Version History;
                1.0 //Boe Prox - 08/20/2016
                    - Initial Version

        .EXAMPLE

            import-csv .\TrendReporting\data.csv|
            out-linechart  -XField Date -YField1 TotalUsedGB -YField2 CapacityGB -Title 'Data Report' `
            -YFieldTitle SizeGB -XFieldIntervalType Months -XFieldInterval 1 -enable3d

        .EXAMPLE

            import-csv .\TrendReporting\data.csv|
            out-linechart  -XField Date -YField1 TotalUsedGB -YField2 CapacityGB -Title 'Data Report' `
            -YFieldTitle SizeGB -XFieldIntervalType Months -XFieldInterval 1 -ToFile -ToFile "C:\users\proxb\desktop\File.jpeg"
    #>
    [cmdletbinding()]
    Param (
        [parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [Object]$InputObject,
        [parameter(Mandatory=$True)]
        [string]$XField,
        [parameter(Mandatory=$True)]
        [string]$YField1,
        [parameter()]
        [string]$YField2,
        [parameter()]
        [string]$Title,
        [parameter()]
        [string]$YFieldTitle,
        [parameter()]
        [string]$YField1Color = 'Red',
        [parameter()]
        $YField2Color = 'Blue',
        [parameter()]
        [int]$XFieldInterval = 16,
        [parameter()]
        [string]$XFieldLabelFormat,
        [parameter()]
        [switch]$Enable3D,
        [parameter()]
        [ValidateSet("Auto", "Number", "Years", "Months", "Weeks", "Days", "Hours", "Minutes", "Seconds", "Milliseconds", "NotSet")]
        [string]$XFieldIntervalType = "Days",
        [parameter()]
        [ValidateSet({
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
        Function ConvertTo-HashTable {
            [cmdletbinding()]
            Param (
                [parameter(ValueFromPipeline=$True)]
                [object]$Object,
                [string[]]$Key
            )
            Begin {
                $HashTable = @{}
            }
            Process {
                ForEach ($Item in $Object) {
                    ForEach ($HKey in $Key) {
                        $HashTable[$Hkey]+=,$Item.$HKey
                    }
                }
            }
            End {
                $HashTable
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
        Function Get-ValueType {
            Param ($Object)
            Switch ($Object) {
                {$Object -as [decimal[]]} {
                    [decimal[]]
                    BREAK
                }
                {$Object -as [datetime[]]} {
                    [datetime[]]
                    BREAK
                }
                {$Object -as [string[]]} {
                    [string[]]
                    BREAK
                }
            }
        }                       
        #endregion Helper Functions
        
        #region Data Processing
        If ($PSBoundParameters.ContainsKey('InputObject')) {
            Write-Verbose "[BEGIN]InputObject bound to parameter name" 
            $ToProcess = $InputObject
        } 
        Else {
            $IsPipeline = $True
            $ToProcess = New-Object System.Collections.ArrayList
            Write-Verbose "[BEGIN]InputObject coming from Pipeline"
        }
        #endregion Data Processing

        #region MSChart Build
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization
        $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
        $Series = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Series
        $ChartTypes = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
        $Series.ChartType = $ChartTypes::Line
        $Chart.Series.Add($Series)
        $Chart.ChartAreas.Add($ChartArea)
        If ($PSBoundParameters.ContainsKey('Enable3D')) {
            $ChartArea.Area3DStyle.Enable3D=$True
            $ChartArea.Area3DStyle.Inclination = 50
        }
        #endregion MSChart Build

        #region MSChart Configuration
        $Chart.Width = 700 
        $Chart.Height = 400 
        $Chart.Left = 10 
        $Chart.Top = 10
        $Chart.BackColor = [System.Drawing.Color]::White

        $ChartTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
        $ChartTitle.Text = $Title
        $Font = New-Object System.Drawing.Font @('Microsoft Sans Serif','12', [System.Drawing.FontStyle]::Bold)
        $ChartTitle.Font =$Font
        $Chart.Titles.Add($ChartTitle)
        $Chart.BorderColor = 'Black'
        $Chart.BorderDashStyle = 'Solid'

        $Chart.ChartAreas.axisy.Title = $YFieldTitle
        $Chart.ChartAreas.axisy.TitleFont = New-Object System.Drawing.Font @('Microsoft Sans Serif','12', [System.Drawing.FontStyle]::Bold)

        $Chart.ChartAreas.axisx.MajorGrid.Enabled=$False
        $Chart.ChartAreas.axisx.MinorTickMark.Enabled=$True
        $Chart.ChartAreas.axisx.MajorTickMark.Enabled=$False

        #endregion MSChart Configuration

        #region Create Legend
        $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
        $Legend.BorderColor = 'Black'
        $Chart.Legends.Add($Legend)
        #endregion Create Legend
    }
    Process {
        #region Pipeline Data Processing
        If ($IsPipeline) {
            Write-Verbose "[PROCESS]Collecting Pipeline InputObject"
            [void]$ToProcess.Add($InputObject)
        } 
        #endregion Pipeline Data Processing
    }
    End {
        Write-Verbose "[END]Processing InputObject"
        $HashTable = $ToProcess | ConvertTo-HashTable -Key $XField,$YField1,$YField2
        If ($PSBoundParameters.ContainsKey('XField')) {
            #Determine Type for conversion of XField
            $XType = Get-ValueType -Object $HashTable[$XField]
            Write-Verbose "XField type is $($Type.fullname)"
            $XFieldData = ($HashTable[$XField] -as $XType)

            $Chart.ChartAreas.axisx.LabelStyle.Format = $XFieldLabelFormat
            $Chart.ChartAreas.axisx.LabelStyle.Angle = -45            
            If ($XType.fullname -match 'datetime') {  
                Write-Verbose "DATETIME"              
                $Chart.ChartAreas.axisx.LabelStyle.IntervalOffsetType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::$XFieldIntervalType
                $Chart.ChartAreas.axisx.LabelStyle.IntervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::$XFieldIntervalType
            } 
            Else {
                $Measured = $HashTable[$XField] | Measure-Object -Minimum -Maximum
                $Chart.ChartAreas.axisx.Maximum = $Measured.Maximum
                $Chart.ChartAreas.axisx.Minimum = $Measured.Minimum            
            }
            $Chart.ChartAreas.axisx.LabelStyle.Interval = $XFieldInterval
        }
        If ($PSBoundParameters.ContainsKey('YField1')) {
            [void]$Chart.Series.Add($YField1) 

            #Determine Type for conversion of YAxis
            $Y1Type = Get-ValueType -Object $HashTable[$YField1]
            Write-Verbose "XField type is $($Type.fullname)"
            $Y1AxisData = ($HashTable[$YField1] -as $Y1Type)

            $Chart.Series[$YField1].Points.DataBindXY($XFieldData, $Y1AxisData)
            $Chart.Series[$YField1].BorderWidth = 2
            $Chart.Series[$YField1].Color = $YField1Color
            $Chart.Series[$YField1].ChartType = 'Line'
        }
        If ($PSBoundParameters.ContainsKey('YField2')) {
            [void]$Chart.Series.Add($YField2) 

            #Determine Type for conversion of YAxis
            $Y2Type = Get-ValueType -Object $HashTable[$YField2]
            Write-Verbose "XField type is $($Type.fullname)"
            $Y2AxisData = ($HashTable[$YField2] -as $Y2Type)

            $Chart.Series[$YField2].Points.DataBindXY($XFieldData, $Y2AxisData)
            $Chart.Series[$YField2].BorderWidth = 2
            $Chart.Series[$YField2].Color = $YField2Color
            $Chart.Series[$YField2].ChartType = 'Line'
        }
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
