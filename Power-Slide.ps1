param(
    $SlideShow = [string]::Empty,
    $Skin = [string]::Empty
)

function List-Slides 
{
    $slides | % { Write-Host $_.Title }

    Read-Host
}

function Insert-Slide ($index)
{
    $newSlide = New-Object PSObject -Property @{
        Title = 'New Slide'
        Subtitle = 'Subtitle'
        Points = 'Your bullet points...'
    }

    if ($index -eq $slides.Length - 1) { $script:slides = $slides + @($newSlide) }
    else { $script:slides = $slides[0..($index)] + @($newSlide) + $slides[($index + 1)..($slides.Length - 1)] }

    Edit-Slide -index ($index + 1)
}

function Remove-Slide ($index)
{
    if ($index -eq 0) { $script:slides = $slides[($index + 1)..($slides.Length - 1)] }
    elseif ($index -eq $slides.Length - 1) { $script:slides = $slides[0..($index - 1)] }
    else { $script:slides = $slides[0..($index - 1)] + $slides[($index + 1)..($slides.Length - 1)] }
}

function Edit-Slides
{
    $selectedItem = 0
    
    do
    {
        cls
        Write-Host 'Power-Slide>Edit slides>' -ForegroundColor Magenta

        Print-Slide $selectedItem
        $key = [Console]::ReadKey().Key.ToString()
        
        switch ($key)
        {
            'PageDown' { if ($selectedItem -lt $slides.Count - 1) { $selectedItem++ } }
            'PageUp'   { if ($selectedItem -gt 0) { $selectedItem-- } }
            'Enter'    { Edit-Slide -index $selectedItem }
            'Insert'   { Insert-Slide -index $selectedItem }
            'Delete'   { 
                if ($slides.Length -gt 1) { Remove-Slide -index $selectedItem }
                if ($selectedItem -eq $slides.Length){ $selectedItem-- }
            }
        }
    } while ($key -ne 'Escape')
}

function Save-Slides
{
    cls
    Write-Host 'Power-Slide>Save slides>' -ForegroundColor Magenta

    $fileName = Read-Host -Prompt 'File name'
    $slides | Export-Csv -Path $fileName
}

function Load-Slides ($fileName = [string]::Empty)
{
    if ($fileName -eq [string]::Empty)
    {
        cls
        Write-Host 'Power-Slide>Load slides>' -ForegroundColor Magenta

        $fileName = Read-Host -Prompt 'File name'
    }

    $Script:slides = Import-Csv -Path $fileName
}

function Print-Welcome
{
cls

Write-Host '    -^^^^^^^^^^^^^----------------------------- -                             ----^^^^/^            '
Write-Host '^////////////^^^^^^^^^^^^^^^^^^^^--------------                               -/c/c^                '
Write-Host '^//////c/////^^^^^^^^^^^^^^^^^^^^^^^^^------------                             ----                 '
Write-Host '^/////cccc////^^^^^^^^^^/^^^^^cua8%%%gau^-----^^^^^^^-------          ------   ----                 '
Write-Host '^////cccc/////////^^^^^^///ca8%8%8axa888gc-----^^/^^^^^^^^^^^---- -u%#######&a^  ---                '
Write-Host '^////cccc//////^^^^^^^^^^^w8aawxucuwwwggaggx/--^^/^^^^^^^^^^--^-/a####&%8gggawc------ --            '
Write-Host '^////ccc/////////^^^^^^^/xwa%gxwuccuwxuuxwwwwa/-^/////u8##%8gga&##&%8awwwxxxuc^------------ ------- '
Write-Host '^///cccc/////////^^^^^^-u8uxu-^^^/uwwwwwwaaxuwwxuxxxa%&88gwawawxxxxxxxxxxxuu/^----------^-----------'
Write-Host '^//wwuaaaggu///^^^^^^^^^gau//^---^uwwwwwaawxuxwa----/awawwxw^-^^////cu///cuu^-----------^^----  - --'
Write-Host '^//cug8%gw%#w//^^^^^^^^/wu^--^^^^-^^cxwwwxxxxxwwc-  -cwxxxxxxgu     ^^/ux/^  --^------------^-----  '
Write-Host '^//ccxww8%%8gx/^/^^^^^^/uc^/^--^^---^/cxxuuuuc/cc/^cwwu///^/ccc///uucuwawwwaaw/-----------------    '
Write-Host '^///ccxaaaaag%w/^//^^^^^xc^^^^^^-^^^uxuuuuuc///^^---^xccxucc^^-----/uxa8%88g%%%&&a/-------------    '
Write-Host '^////ccccxagg%##&a///^^^cxc^^^^---^-^uxxuuuc^--/^u//ua%uuu^^---^/////cxwg%&%gagg8%&&a^--------      '
Write-Host '^//////ccccuwaaag%&##8wuuuuc^^^^-----/uxxu^^^//^/8&%w//uxcuxc///^^///uucuwa88g888888%&8c-------  -  '
Write-Host '/c///////cccccxwaaggg%&###8wuc^------/uu///////^^wwuuwwuxxu/-^^^-^cc^/uw//xcxu/uwaawxc^^cu/-------  '
Write-Host '/ccc///////////uxxwagggggg8%&8xc/xawxxxucccuc//uc/^---------^c/----^^-cxxwu//^^^^///cccc^-^/^------ '
Write-Host '^c//////cccccc//cxwwwaaggggggg%###&gxc///c/^^/^-^ccuuc/^^^^^^^--^^--^^----^^^-^^^^^^/^^^^^/cax----- '
Write-Host '/cc//////////////^/xwxwwaagggaagggg8gu^^//^^cc/^^^^^^^^---^^^---^^^^^^^^^^^^--^--^^^^^^^/uxxwww/--  '
Write-Host '/ccc////////////^^/^/xggawwaaaawwwwwawc---^^^^^^^^^^^^-^^^^^^^^^^^^^^^^-^^^^-^^^^---^--- -/wgc/%%%gc'
Write-Host '/cc////////////^^^^//^^/uxa8gawwwwwwwwx^-----^^^-^^^^^^^^^^^^^^^^^^^---^^^^-^^-------^////cwc^/x/^gw'
Write-Host '/cc//////^^^^^^^^^^^^^^^^^^^^^^---^cxxwc--------^--^^^^^^^^^^^^^^^^^^^^^---^------^^^^^//x8u^/c//^wa'
Write-Host '/cc//////^^^^^^^^^^^^^^///////^^^^---^^-------^^^^^^/^^^^^^^^^^^^^^^^^^^^-^^--------^^^^cwc^uu////cw'
Write-Host '/c////////^^^^^^^^^^^^^^///////^^-^^^^--------^^^^//^//^^^^^^^^^^^^^^^^^^^-^^^--------^^ccxaa8%%guuc'
Write-Host '^////^^^^^-------^^^^^^^^^^/////^^^^^^/cuuuc///^^-^x/ccc^^^^^^^^^^^^^^^^^^^-^--------------^cwg8%&%c'
Write-Host '-^^^^^^^^^-----^-^^^^^^^^^^////ccccccccccccccuuuc/^^-/uxxc^^^^^^^^^^^^^^^^^---------------^//wagg8&a'
Write-Host '-^^^^-^-------^^^/ccccuuuuuuuuxxxxxuuuuuxxxxxxxxxuuuuuucuxxu^--^^^^^^---^------^---------^^/xx//xwax'
Write-Host '------^^^//ccccccccccuccuuuuuuuuuuuuuuuuuuuuxxxuuuuccccucua8gwu/^-^^^^^^^----^---------^^//^^//cwwwu'
Write-Host '-^^^/cccccccccuuuuuuuuuuuuuccuucuuuuuuuuuxxxxxxuuuuuccccu88awxuc^-----^^--------------^^^^^///^uccuc'
Write-Host '-/cuucccuuuuuxxxxxuuuuxxuuuuuuuxxxxxxxxxxxxxxuuuucccuccccxxxc^- -^/cuuuucc/^ --^  ---^^^^^^^/^/ccc/^'
Write-Host '^uuuuuuuuuxwggg8gawxxxuuxxxxxxxuuuuxxxxxxxxxxxxxxxuuuuuuuucuu//ccuuxuuc//ccc/^^^-------^^^^^^^/^///^'
Write-Host '/cccccccuuuuuuuuuuuuuuuuuuuuuxuxxxxxwxxxxxxxwxxxxuuuuuuucc/ugw/uuuuucug%8u///////^^^^^^^^^^^^^^///c/'
Write-Host '^cccccccccccccccuuuuuuucuccuccuccuuuuxxxxxxxxxxxuuuuuuuuucuucccuccc/a%%8u//^//^//^^^^^^^^^^^///cc/^-'
Write-Host '^uuccccccccccccccccccccccccccccuxuuuuuuuwg88gawucucuuuuucccccccc/c/u8%%wccc/^^^^^^^^/^^^^^^^////^^^^'
Write-Host '-//c/////cucccccuccccccucccuucuxxuxxxuc/////^^^^-^^^u/^^^-----^^/cc/a8%w//cc///^^^^^^^^^^^^^^^////^-'
Write-Host '-^^/cccccccccccccccccccuxxawxxucuuuuuc//^^///^^^---------------  --^//cc///cc////^^^^^^^^^^^^^^^///^'
Write-Host '-^-^^//ccccccccccccccccccccccccccccuuuc//////^^^---^----------------/xx-  --^^/////////^^^^^^^^^/uu^'
Write-Host ''
Write-Host -ForegroundColor Cyan 'Welcome to Power-Slide v0.0.1'
Write-Host ''
}

function Print-Slide ($index)
{
    $slide = $slides[$index]
    icm $theme -ArgumentList $slide
}

function Edit-Slide($index)
{
    do
    {
        cls
        Write-Host "Power-Slide>Edit slides>$index>" -ForegroundColor Magenta

        Print-Slide $index
        $token = Read-Host
        
        switch ($token)
        {
            ':title' { $slides[$index].Title = Read-Host -Prompt 'Title' }
            ':bullet' { $slides[$index].Points = Read-Host -Prompt 'Bullet points ( default separator: @)' }
            ':subtitle' { $slides[$index].Subtitle = Read-Host -Prompt 'Subtitle' }
        }
    } while ($token -ne ':done')
}

function Load-Theme ($fileName = [string]::Empty)
{
    if ($fileName -eq [string]::Empty)
    {
        cls
        Write-Host 'Power-Slide>Load theme>' -ForegroundColor Magenta

        $filename = Read-Host -Prompt 'Path to theme file'
        
        if ((-not $fileName.Contains('\')) -or (-not $fileName.Contains('/'))) { $fileName = ".\$fileName" }
    }
    
    if (-not (Test-Path $filename)) 
    {
        $invocation = Split-Path -Parent $PSCommandPath
        $filename = "$invocation\$filename"
    }

    if (Test-Path $filename) { $script:theme = . $filename }

    else { Push-Error 'Could not find theme file' }
}

function Push-Error ($message)
{
    $errorStack.Push($message)
}

function Pop-Error
{
    while ($errorStack.Count -ne 0)
    {
        Write-Host "Error: $($errorStack.Pop())" -ForegroundColor Red
        Write-Host ''
    }
}

function Sanitize ($input)
{
    
}

function global:next
{
    if ($global:_selectedItem -lt $global:_slides.count - 1) 
    {
        $global:_selectedItem++

        print
    }
}

function global:prev
{
    if ($global:_selectedItem -gt 0) 
    {
        $global:_selectedItem--

        print
    }
}

function global:print
{
    cls
    Invoke-Command $global:_theme -ArgumentList $global:_slides[$global:_selectedItem]
}

function global:done
{
    Remove-Item Function:\next
    Remove-Item Function:\prev
    Remove-Item Function:\print
    Remove-Item Function:\done

    Remove-Variable -Name '_slides' -Scope Global
    Remove-Variable -Name '_theme' -Scope Global
    Remove-Variable -Name '_selectedItem' -Scope Global
}

$defaultTheme = {
    param($slide)
    
    $title = $slide.Title
    $subtitle = $slide.Subtitle
    $points = if ($slide.Points.Length -ne 0) {$slide.Points.Split('@')} else {@()}
    $other = $slide.Other

    $slideWidth = 40
    $slideHeight = 16

    $prefixLength = [Math]::Floor(($slideWidth - $title.Length) / 2) - 1
    $suffixLength = $slideWidth - $prefixLength - $title.Length - 2
       
    
    Write-Host ('-' * $slideWidth) -ForegroundColor Cyan

    if ($points.Count -ne 0)
    {
        Write-Host ("|$(' ' * $prefixLength)$($title)$(' ' * $suffixLength)|") -ForegroundColor Cyan
        Write-Host ('-' * $slideWidth) -ForegroundColor Cyan
        Write-Host ''
        Write-Host ''

        $rowsLeft = $slideHeight - 5
        $points | %{
            Write-Host '> ' -NoNewline

            if ($_.Length -gt $slideWidth) {
                for($i = 0; $i -lt ($_.Length/$slideWidth); $i++) {
                    $start = $i * $slideWidth
                    $length = $slideWidth,($_.Length - $start)
                    Write-Host ($_.Substring($start, $length))
                }
            } else {
                Write-Host $_
            }

            $rowsLeft--
        }

        Write-Host ("`n" * ($rowsLeft - 1))
    }
    else
    {
        $subPrefixLength = [Math]::Floor(($slideWidth - $subTitle.Length) / 2) - 1
        $subSuffixLength = $slideWidth - $subPrefixLength - $subTitle.Length - 2

        Write-Host ("`n" * ([Math]::Floor($slideHeight/2) - 4))
        Write-Host ('-' * $slideWidth) -ForegroundColor Cyan
        Write-Host ("|$(' ' * $prefixLength)$($title)$(' ' * $suffixLength)|") -ForegroundColor Cyan
        Write-Host ("|$(' ' * ([Math]::Floor($slideWidth/2) - 2))--$(' ' * ([Math]::Floor($slideWidth/2) - 2))|") -ForegroundColor Cyan
        Write-Host ("|$(' ' * $subPrefixLength)$($subTitle)$(' ' * $subSuffixLength)|") -ForegroundColor Cyan
        Write-Host ('-' * $slideWidth) -ForegroundColor Cyan
        Write-Host ("`n" * ([Math]::Floor($slideHeight/2) - 4))
    }
    
    Write-Host ('-' * $slideWidth) -ForegroundColor Cyan
}

$token = ""

$slide1 = New-Object PSObject -Property @{
    Title = 'New Slide'    
    Points = 'Point 1@Point 2@Point 3'
    Subtitle = 'Subtitle'
}

$slides = @($slide1)
$globalIndex = 0
$theme = $defaultTheme
$errorStack = New-Object System.Collections.Stack

if ($SlideShow -ne [string]::Empty) { Load-Slides $SlideShow }
if ($Skin -ne [string]::Empty) { Load-Theme $Skin }

while ($token -ne ":q")
{
    Print-Welcome
    Pop-Error
    Write-Host 'Power-Slide>' -ForegroundColor Magenta
    $token = Read-Host
    
    try
    {
        switch ($token)
        {
            ':list' { List-Slides }
            ':edit' { Edit-Slides }
            ':save' { Save-Slides }
            ':theme'{ Load-Theme  }
            ':load' { Load-Slides }
        
            ':present' 
            {
                $token = ':q'

                New-Variable -Name '_slides' -Value $slides -Scope Global -Force
                New-Variable -Name '_theme' -Value $theme -Scope Global -Force
                New-Variable -Name '_selectedItem' -Value (-1) -Scope Global -Force
            }
        }
    }
    catch
    {
        Push-Error $_.Exception.Message
    }
}
