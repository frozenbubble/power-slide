{
    param($slide)
    
    $title = $slide.Title
    $subtitle = $slide.SubTitle
    $points = $slide.Points.Split('@')
    $other = $slide.Other

    $slideWidth = 40
    $slideHeight = 16

    $prefixLength = [Math]::Floor(($slideWidth - $title.Length) / 2) - 1
    $suffixLength = $slideWidth - $prefixLength - $title.Length - 2
        
    Write-Host ('-' * $slideWidth) -ForegroundColor Green
    Write-Host ("|$(' ' * $prefixLength)$($title)$(' ' * $suffixLength)|") -ForegroundColor White
    Write-Host ('-' * $slideWidth) -ForegroundColor Green
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
    Write-Host ('-' * $slideWidth) -ForegroundColor Green
}