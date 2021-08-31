# PowerShell port of "The Virtual Piano" piano.bas originally by written by Chris Peters 1983
# Included with Microsoft Mouse https://winworldpc.com/product/microsoft-mouse/1x

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form

# zooms the display from the original 320x200
$global:multiplier = 4

# Key Size Parameters
$global:YL = 60 
$global:WKL = 80 
$global:BKL = 45 
$global:KW = 15
$global:WKN = 21
$global:XL = 320-$global:KW*$global:WKN
$global:YH = $global:YL + $global:WKL
$global:XH = 319
$global:BKW2=$global:KW / 3
$global:QX = 272 
$global:QY = 176

$form.add_MouseDown({
    param($sender,$e)
    
    [Int]$MX = $e.X # mouse X position
    [Int]$MY = $e.Y # mouse Y position

    $WKY = [Math]::Ceiling(($MX-($global:XL*$global:multiplier))/($global:KW * $global:multiplier))
    $R = 1  
    if ($MY -le (($global:YL+$global:BKL)*$global:multiplier))
    {
        $MK=($MX-($global:XL*$global:multiplier)) % ($global:KW * $global:multiplier)
        if ($MK -le ($global:BKW2*$global:multiplier))
        { 
            $R = 0
        }
        elseif ($MK -ge (($global:KW-$global:BKW2)*$global:multiplier))
        {
            $R = 2  
        }
    }

    $global:samplePlayer[$global:Freq[$WKY,$R]].Play()
})

$form.add_MouseUp({
    # if caps lock is on won't stop the sound on mouse up
    if (!([System.Windows.Forms.Control]::IsKeyLocked([System.Windows.Forms.Keys]::CapsLock)))
    {
        ForEach ($sample in $global:samplePlayer)
        {
            $sample.Stop()
        }
    }
})

$form.BackColor = [System.Drawing.Color]::FromArgb(0x00,0x00,0xAA)
$form.Width = 320 * $global:multiplier
$form.Height = 200 * $global:multiplier
$form.Visible = $true 
$logoBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0xAA,0xAA,0xAA))
$whiteKeyBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0xAA,0xAA,0xAA))
$blackKeyBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0xAA,0x00,0xAA))

$backgroundBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0x00,0x00,0xAA))
$gfx = $form.CreateGraphics()

# Load Piano Samples
$pianoSamples = @( 
    "Piano.mf.C3.wav",
    "Piano.mf.Db3.wav",
    "Piano.mf.D3.wav",
    "Piano.mf.Eb3.wav",
    "Piano.mf.E3.wav",
    "Piano.mf.F3.wav",
    "Piano.mf.Gb3.wav",
    "Piano.mf.G3.wav",
    "Piano.mf.Ab3.wav",
    "Piano.mf.A3.wav",
    "Piano.mf.Bb3.wav"
    "Piano.mf.B3.wav",
    "Piano.mf.C4.wav",
    "Piano.mf.Db4.wav",
    "Piano.mf.D4.wav",
    "Piano.mf.Eb4.wav",
    "Piano.mf.E4.wav",
    "Piano.mf.F4.wav",
    "Piano.mf.Gb4.wav",
    "Piano.mf.G4.wav",
    "Piano.mf.Ab4.wav",
    "Piano.mf.A4.wav",
    "Piano.mf.Bb4.wav"
    "Piano.mf.B4.wav",
    "Piano.mf.C5.wav",
    "Piano.mf.Db5.wav",
    "Piano.mf.D5.wav",
    "Piano.mf.Eb5.wav",
    "Piano.mf.E5.wav",
    "Piano.mf.F5.wav",
    "Piano.mf.Gb5.wav",
    "Piano.mf.G5.wav",
    "Piano.mf.Ab5.wav",
    "Piano.mf.A5.wav",
    "Piano.mf.Bb5.wav"
    "Piano.mf.B5.wav")

$global:samplePlayer = @()   
ForEach ($filename in $pianoSamples)
{
    $path = Join-Path -Path $PSScriptRoot -childPath $filename
    $sample = New-Object System.Media.SoundPlayer($path)
    $sample.Load()
    $global:samplePlayer += $sample 
} 

$global:Freq = New-Object 'Int[,]' 22,3
$j = 0
For ($i=0;$i -lt 21;$i+=7)
{
    $global:Freq[(1+$i),1]=0+$j
    $global:Freq[(1+$i),2]=1+$j
    $global:Freq[(2+$i),0]=1+$j
    $global:Freq[(2+$i),1]=2+$j
    $global:Freq[(2+$i),2]=3+$j
    $global:Freq[(3+$i),0]=3+$j
    $global:Freq[(3+$i),1]=4+$j
    $global:Freq[(3+$i),2]=4+$j
    $global:Freq[(4+$i),0]=5+$j
    $global:Freq[(4+$i),1]=5+$j
    $global:Freq[(4+$i),2]=6+$j
    $global:Freq[(5+$i),0]=6+$j
    $global:Freq[(5+$i),1]=7+$j
    $global:Freq[(5+$i),2]=8+$j
    $global:Freq[(6+$i),0]=8+$j
    $global:Freq[(6+$i),1]=9+$j
    $global:Freq[(6+$i),2]=10+$j
    $global:Freq[(7+$i),0]=10+$j
    $global:Freq[(7+$i),1]=11+$j
    $global:Freq[(7+$i),2]=11+$j
    $j+=12
}

# Draw old school Microsoft Logo from Piano.bas, included with Microsoft Mouse for MS-DOS
$msLogo = @( 208, 28 )
$msLogo += @( 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
$msLogo += @( 0, 0, 0, 0, 0, 0, -128, 0, 0, 0, 0, 0, 0)
$msLogo += @( 0, 0, 3, -512, 0, 7, -16, 62, 0, -128, 0, 0, 0)
$msLogo += @( -512, 4071, -16369, -128, 8188, 0, 0, 255, -32765, -32, 1023, -7169, -128)
$msLogo += @( -256, 8167, -16321, -32, 8191, -32705, -2, 1023, -8177, -8, 1023, -7169, -128)
$msLogo += @( -256, 8167, -16129, -8, 8191, -16260, 31, 2047, -4033, -2, 1023, -7169, -128)
$msLogo += @( -128, 16359, -15880, 252, 8191, -8192, 0, 2047, -3970, 63, 1023, -7169, -128)
$msLogo += @( -128, 16359, -15392, 62, 7943, -3585, -1, -14361, -3848, 15, -31745, -7169, -128)
$msLogo += @( -64, 32743, -15424, 30, 7939, -3136, 1, -6173, -3856, 7, -31776, 7, -16384)
$msLogo += @( -64, 32743, -14464, 15, 7937, -4096, 0, 2016, 480, 3, -15392, 7, -16384)
$msLogo += @( -32, -25, -14464, 15, 7937, -2049, 127, -2056, 480, 3, -15392, 7, -16384)
$msLogo += @( -32, -25, -12544, 7, -24829, -2176, 0, -3074, 960, 1, -7200, 7, -16384)
$msLogo += @( -15, -25, -12544, 0, 7943, -4096, 0, 1023, -31808, 1, -7169, -8185, -16384)
$msLogo += @( -1039, -1049, -12544, 0, 8191, -6146, 63, -3585, -15424, 1, -7169, -8185, -16384)
$msLogo += @( -1029, -1049, -12544, 0, 8191, -6272, 0, -3969, -7232, 1, -7169, -8185, -16384)
$msLogo += @( -1541, -3097, -12544, 0, 8191, -16384, 0, 15, -3136, 1, -7169, -8185, -16384)
$msLogo += @( -1537, -3097, -12544, 0, 8191, -14464, 0, -4093, -1088, 1, -7200, 7, -16384)
$msLogo += @( -1793, -7193, -14464, 0, 7967, -14337, 127, -4095, -1568, 3, -15392, 7, -16384)
$msLogo += @( -1793, -7193, -14464, 15, 7943, -8192, 0, 4033, -1568, 3, -15392, 7, -16384)
$msLogo += @( -1921, -15385, -15424, 30, 7939, -7232, 1, -4125, -1808, 7, -31776, 7, -16384)
$msLogo += @( -1921, -15385, -15392, 62, 7939, -3585, -1, -12289, -1800, 15, -31776, 7, -16384)
$msLogo += @( -1985, -31769, -15880, 252, 7937, -4096, 0, 2047, -3970, 63, 992, 7, -16384)
$msLogo += @( -1985, -31769, -16130, 1016, 7937, -3972, 31, 2047, -8129, -32514, 992, 7, -16384)
$msLogo += @( -2017, 999, -16321, -32, 7937, -4033, -2, 1023, -16369, -8, 992, 7, -16384)
$msLogo += @( -2017, 999, -16369, -128, 7937, -4096, 0, 255, 3, -32, 992, 7, -16384)
$msLogo += @( 0, 0, 3, -512, 0, 7, -16, 0, 0, -128, 0, 0, 0)
$msLogo += @( 0, 0, 0, 0, 0, 0, -128, 0, 0, 0, 0, 0, 0)
$msLogo += @( 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

$width = $msLogo[0]
$height = $mslogo[1]
$c=0
$currentRow = 0
$currentColumn = 0
$totalColors = @()

$total = ""
$offSet = 63
for ($i = 2;$i -lt $mslogo.Length ;$i++)
{
    for ($j =16;$j -gt 0;$j--)
    {
        if (([Int16]$mslogo[$i] -band (1 -shL $j)) -ne 0)
        {
            $gfx.FillRectangle($logoBrush,($offSet*$global:multiplier)+$currentColumn*$global:multiplier,$currentRow*$global:multiplier,$global:multiplier,$global:multiplier)

        }
        $currentColumn++
        if ($currentColumn -ge $WIDTH)
        {
            $currentColumn = 0 
            $currentRow++
        }   
    }
}

# Draw white keys
$gfx.FillRectangle($whiteKeyBrush,$global:XL*$global:multiplier,$global:YL*$global:multiplier,($global:XH*$global:multiplier)-($global:XL*$global:multiplier),($global:YH*$global:multiplier)-($global:YL*$global:multiplier))
for ($i=$global:XL;$i -lt $global:XH;$i+=$global:KW)
{
    $gfx.FillRectangle($backgroundBrush,$i*$global:multiplier,$global:YL*$global:multiplier,$global:multiplier,($global:YH*$global:multiplier)-($global:YL*$multipler))
}

# Draw black keys
$C=6
for ($X=$global:XL;$X -lt $global:XH;$X+=$global:KW)
{
    $C++
    if ($C -eq 7) { $C = 0 }
    if (!($C -eq 0 -or $C -eq 3))
    {
         $gfx.FillRectangle($blackKeyBrush,($X-$global:BKW2)*$global:multiplier,$global:YL*$global:multiplier,(($X+$global:BKW2)*$global:multiplier)-(($X-$global:BKW2)*$global:multiplier),(($global:YL+$global:BKL)*$global:multiplier)-($global:YL*$global:multiplier))
    }
}

[void][System.Windows.Forms.Application]::Run($form)
$Form.close()

