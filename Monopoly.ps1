# Clear all non-constant variables
Remove-Variable * -ErrorAction SilentlyContinue

# Stop game on error
trap {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host "Character: $($_.InvocationInfo.OffsetInLine)" -ForegroundColor Red
    if ($Window) {
        $Window.TaskbarItemInfo.Overlay = $null
        $Window.Dispatcher.Invoke([Action]{$Window.Close()})
    }
    Exit 1
}

# Some settings
$host.UI.RawUI.WindowTitle = "Monopoly!"
$Testing = $false
$Pause = 1000 # In Milliseconds
$TotalHouses = 0
$TotalHotels = 0

# Load some assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Set the stage
[Console]::BackgroundColor = 'Black'
Clear-Host

# Title
Write-Host -ForegroundColor Red @'
$$\      $$\                                                   $$\           
$$$\    $$$ |                                                  $$ |          
$$$$\  $$$$ | $$$$$$\  $$$$$$$\   $$$$$$\   $$$$$$\   $$$$$$\  $$ |$$\   $$\ 
$$\$$\$$ $$ |$$  __$$\ $$  __$$\ $$  __$$\ $$  __$$\ $$  __$$\ $$ |$$ |  $$ |
$$ \$$$  $$ |$$ /  $$ |$$ |  $$ |$$ /  $$ |$$ /  $$ |$$ /  $$ |$$ |$$ |  $$ |
$$ |\$  /$$ |$$ |  $$ |$$ |  $$ |$$ |  $$ |$$ |  $$ |$$ |  $$ |$$ |$$ |  $$ |
$$ | \_/ $$ |\$$$$$$  |$$ |  $$ |\$$$$$$  |$$$$$$$  |\$$$$$$  |$$ |\$$$$$$$ |
\__|     \__| \______/ \__|  \__| \______/ $$  ____/  \______/ \__| \____$$ |
                                           $$ |                    $$\   $$ |
                                           $$ |                    \$$$$$$  |
                                           \__|                     \______/ 
By Ray Smalley

'@

# Start game
if (!$psISE) {
    Write-Host "Press SPACEBAR to play!"`n
    do {
        while ([System.Console]::KeyAvailable) {
            [System.Console]::ReadKey($true) | Out-Null
        }
        $KeyInfo = [System.Console]::ReadKey($true)
        $KeyChar = $KeyInfo.KeyChar
    } while ($KeyChar -ne ' ')
}

# Dice rolling function
function DiceRoll {
    $DiceFaces = "$PSScriptRoot\Dice.png"
    $DiceFacesImage = [System.Windows.Media.Imaging.BitmapImage]::New($DiceFaces)

    # Get individual dice faces from image
    function GetDiceFace($FaceNumber) {
        $DiceX = ($FaceNumber - 1) * 200
        $DiceY = 0
        $DiceRect = [System.Windows.Int32Rect]::New($DiceX, $DiceY, 200, 200)
        return [System.Windows.Media.Imaging.CroppedBitmap]::New($DiceFacesImage, $DiceRect)
    }

    $DiceImages = @( $null, (GetDiceFace 1), (GetDiceFace 2), (GetDiceFace 3), (GetDiceFace 4), (GetDiceFace 5), (GetDiceFace 6))
    $LastDie1Roll = $null
    $LastDie2Roll = $null

    # Simple dice roll animation
    for ($i = 0; $i -lt 10; $i++) {
        do {
            $Die1Roll = (Get-Random -Minimum 1 -Maximum 7)
        } while ($Die1Roll -eq $LastDie1Roll)
        $LastDie1Roll = $Die1Roll

        do {
            $Die2Roll = (Get-Random -Minimum 1 -Maximum 7)
        } while ($Die2Roll -eq $LastDie2Roll)
        $LastDie2Roll = $Die2Roll

        # Update dice images in the UI
        $DieImage1.Source = $DiceImages[$Die1Roll]
        $DieImage2.Source = $DiceImages[$Die2Roll]
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        
        if (!$Testing) {Start-Sleep -Milliseconds (100 + ($i * 10))}
    }

    if ($Die1Roll -eq $Die2Roll) {$Double = $true} else {$Double = $false}
    $Roll = $Die1Roll + $Die2Roll
    return $Roll, $Double
}

# Get key press function
function GetKeyPress {
    if (!$psISE) {
        while ([System.Console]::KeyAvailable) {
            [System.Console]::ReadKey($true) | Out-Null
        }
        $KeyInfo = [System.Console]::ReadKey($true)
        $KeyChar = $KeyInfo.KeyChar

        if ($KeyChar -eq 'y') {
            $Answer = $true
            Write-Host
        } elseif ($KeyChar -eq 'n') {
            $Answer = $false
            Write-Host
        } else {
            GetKeyPress
        }
        return $Answer
    } else {
        return (Read-Host)
    }
}

# Advance to nearest... function
function NextDestination {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$Name
    )
    $NextSpaceIndex = ($Player.CurrentSpace + 1) % $Spaces.Count
    while ($NextDestination.Type -notmatch $Name) {
        $NextDestination = $Spaces[$NextSpaceIndex]
        $NextSpaceIndex = ($NextSpaceIndex + 1) % $Spaces.Count
    }
    return $NextDestination.Number
}

# Property color function
function GetPropertyColor($SetColor) {
    switch ($SetColor) {
        'Black' {'DarkGray'}
        'Blue' {'Blue'}
        'Cyan' {'Cyan'}
        'Green' {'Green'}
        'LightBlue' {'Cyan'}
        'Magenta' {'Magenta'}
        'Orange' {if ($psISE) {'DarkYellow'} else {[char]27 + "[38;2;255;165;0m"}}
        'Purple' {'DarkMagenta'}
        'Red' {'Red'}
        'Yellow' {'Yellow'}
        default {'White'}
    }
}
$ResetColor = [char]27 + "[0m"

# Set available player colors
$Colors = @(
'Blue'
'Cyan'
'Green'
'DarkCyan'
'DarkMagenta'
'Magenta'
'Red'
'Yellow'
)
$Colors = $Colors | Get-Random -Count $Colors.Count

# Prompt the user for the total number of players
$NumPlayers = 0
do {
    $UserInput = Read-Host "Enter the number of players"
    $IsNumber = [int]::TryParse($UserInput, [ref]$NumPlayers)
    if (!$IsNumber) {
        Write-Host "Please enter a valid number!"`n -ForegroundColor Red
    } elseif ($NumPlayers -lt 2 -or $NumPlayers -gt 8) {
        Write-Host "Please choose a number between 2 and 8"`n -ForegroundColor Red
    }
} until ($IsNumber -and $NumPlayers -ge 2 -and $NumPlayers -le 8)
Write-Host

# Prompt user for number of AI players
$NumAIPlayers = 0
do {
    $UserInput = Read-Host "Enter the number of AI players"
    $IsNumber = [int]::TryParse($UserInput, [ref]$NumAIPlayers)
    if (!$IsNumber) {
        Write-Host "Please enter a valid number!"`n -ForegroundColor Red
    } elseif ($NumAIPlayers -lt 0 -or $NumAIPlayers -gt $NumPlayers) {
        Write-Host "Please choose a number between 0 and $NumPlayers"`n -ForegroundColor Red
    }
} until ($IsNumber -and $NumAIPlayers -ge 0 -and $NumAIPlayers -le $NumPlayers)
Write-Host

# Check if this an AI-only game
if ($NumPlayers -eq $NumAIPlayers) {$AIOnlyGame = $true}

# Set player information
$Players = @()
for ($i = 1; $i -le ($NumPlayers - $NumAIPlayers); $i++) {
    New-Variable -Name Player$i
    Get-Variable -Name Player$i | Set-Variable -Value (New-Object -TypeName psobject -Property @{
        Name = Read-Host "Enter the name of Player $i"
        IsHuman = $true
        Money = 1500
        Color = $Colors[$i-1]
        CurrentSpace = 0
        DoubleCount = 0
        InJail = $false
        GetOutOfJailFree = 0
        HousesOwned = $null
        HotelsOwned = $null
    })
    $Players += (Get-Variable -Name Player$i -ValueOnly)
    Write-Host
}

# AI players
$Names = @("Josh", "Beverly", "Bob", "Nikki", "Ray", "Theresa", "Oliver", "Eva", "Freddie", "Cow")
$Names = $Names | Where {$Players.Name -notcontains $_} | Get-Random -Count $Names.Count

# Set AI player information
for ($i = $i; $i -le $NumPlayers; $i++) {
    New-Variable -Name Player$i
    Get-Variable -Name Player$i | Set-Variable -Value (New-Object -TypeName psobject -Property @{
        Name = $Names[$i-1]
        IsHuman = $false
        Money = 1500
        Color = $Colors[$i-1]
        CurrentSpace = 0
        DoubleCount = 0
        InJail = $false
        GetOutOfJailFree = 0
        HousesOwned = $null
        HotelsOwned = $null
    })
    $Players += (Get-Variable -Name Player$i -ValueOnly)
}

# House/hotel selling function
function SellBuildings {
    while ($Player.Money -le 0) {
        $PropertiesWithBuildings = $Spaces | Where {$_.Owner -eq $Player.Name -and ($_.HotelCount -gt 0 -or $_.HouseCount -gt 0)}
        $PropertiesSetsWithBuildings = $PropertiesWithBuildings | Group-Object {$_.SetColor}
        $NoPropertiesToSell = $true
        foreach ($ColorSet in $PropertiesSetsWithBuildings) {
            $SortedProperties = $ColorSet.Group | Sort-Object -Property @{Expression = {($_.HotelCount * 5) + $_.HouseCount}; Descending = $true}
            $Property = $SortedProperties | Select -First 1
            if ($Player.Money -le 0) {
                $BuildingSale = $Property.BuildingCost / 2
                if ($Property.HotelCount -gt 0) {
                    # Sell hotel
                    $SpaceColor = GetPropertyColor $Property.SetColor
                    Write-Host "$($Player.Name) sold a hotel on " -NoNewline -ForegroundColor $Player.Color
                    if ($Property.SetColor -eq 'Orange' -and !$psISE) {
                        Write-Host $SpaceColor"$($Property.Name)"$ResetColor -NoNewline
                    } else {
                        Write-Host "$($Property.Name)" -NoNewline -ForegroundColor $SpaceColor
                    }
                    Write-Host " for `$$BuildingSale"`n -ForegroundColor $Player.Color
                    $Property.HotelCount = 0
                    $Property.HouseCount = 4
                    $Player.HotelsOwned -= 1
                    $Player.HousesOwned += 4
                    $TotalHotels--
                    $Player.Money += $BuildingSale
                    UpdateProperties
                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    $NoPropertiesToSell = $false
                } elseif ($Property.HouseCount -gt 0) {
                    # Sell house
                    $SpaceColor = GetPropertyColor $Property.SetColor
                    Write-Host "$($Player.Name) sold a house on " -NoNewline -ForegroundColor $Player.Color
                    if ($Property.SetColor -eq 'Orange' -and !$psISE) {
                        Write-Host $SpaceColor"$($Property.Name)"$ResetColor -NoNewline
                    } else {
                        Write-Host "$($Property.Name)" -NoNewline -ForegroundColor $SpaceColor
                    }
                    Write-Host " for `$$BuildingSale"`n -ForegroundColor $Player.Color
                    $Property.HouseCount--
                    $Player.HousesOwned--
                    $TotalHouses--
                    $Player.Money += $BuildingSale
                    UpdateProperties
                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    $NoPropertiesToSell = $false
                }
            }
        }

        if ($NoPropertiesToSell) {
            break
        }
    }
}

# Mortgaging properties function
function MortgageProperties {
    $PropertiesOwned = $Spaces | Where {$_.Owner -eq $Player.Name}
    $PropertiesOwned | Where {!$_.Mortgaged -and $_.HouseCount -eq 0 -and $_.HotelCount -eq 0} | ForEach-Object {
        if ($Player.Money -le 0) {
            $SpaceColor = GetPropertyColor $_.SetColor
            Write-Host "$($Player.Name) mortgaged " -NoNewline -ForegroundColor $Player.Color
            if ($_.SetColor -eq 'Orange' -and !$psISE) {
                Write-Host $SpaceColor"$($_.Name)"$ResetColor -NoNewline
            } else {
                Write-Host "$($_.Name)" -NoNewline -ForegroundColor $SpaceColor
            }
            Write-Host " for `$$($_.Mortgage)"`n -ForegroundColor $Player.Color
            $_.Mortgaged = $true
            $Player.Money += $_.Mortgage
            UpdateProperties
            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
        } else {
            return # Stop once the player is solvent
        }
    }
}

# Bankruptcy check function
function BankruptcyCheck {
    $RemainingPlayers = @()
    foreach ($Player in $Players) {
        # First try to sell buildings
        if ($Player.Money -le 0 -and ($Player.HousesOwned -gt 0 -or $Player.HotelsOwned -gt 0)) {
            if ($Player.IsHuman) {
                Write-Host "You are indebted `$$($Player.Money). Your buildings will be sold until you are solvent."
                Read-Host
            }
            SellBuildings
        }
        # Then mortgage properties
        if ($Player.Money -le 0) {
            if ($Player.IsHuman) {
                Write-Host "You are indebted `$$($Player.Money). Your properties will be mortgaged until you are solvent."
                Read-Host
            }
            MortgageProperties
        }
        # If no assets to liquify...
        if ($Player.Money -le 0) {
            $IsBankrupt = $true
            $Spaces | ForEach-Object {
                if ($_.Owner -eq $Player.Name) {
                    $_.Owner = 'Bank'
                    $_.Mortgaged = $false
                }
            }
            UpdatePropertyOwner
            UpdateProperties

            # Remove player from board
            $PlayerGrid = $PlayerGrids[$Player.Name]
            if ($PlayerGrid -ne $null) {
                $BoardCanvas = $Window.FindName('BoardCanvas')
                $BoardCanvas.Children.Remove($PlayerGrid)
            }
							  
            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
        } else {
            $RemainingPlayers += $Player
            $IsBankrupt = $false
        }
        UpdatePlayerPieces
    }
    return @{
        RemainingPlayers = $RemainingPlayers
        IsBankrupt = $IsBankrupt
        HumanPlayersLeft = ($Players | Where-Object {$_.IsHuman -eq $true} | Measure-Object).Count
    }
}

# Chance Cards
$ChanceCards = @(
    @{
        Name = 'Take a walk on the Boardwalk'
        Space = 39
    }
    @{
        Name = 'Advance to Go'
        Space = 0
    }
    @{
        Name = 'Advance to Illinois Avenue'
        Space = 24
    }
    @{
        Name = 'Advance to St. Charles Place'
        Space = 11
    }
    @{
        Name = 'Advance to the nearest Railroad'
    }
    @{
        Name = 'Advance to the nearest Utility'
    }
    @{
        Name = 'Bank pays you dividend of $50!'
        Money = 50
    }
    @{
        Name = 'Get Out of Jail Free'
    }
    @{
        Name = 'Go Back 3 Spaces'
        SpaceMove = -3
    }
    @{
        Name = 'Go to Jail'
        Space = 10
        InJail = $true
    }
    @{
        Name = 'Make general repairs on all your property. For each house pay $25. For each hotel pay $100.'
        MoneyPerHouse = 20
        MoneyPerHotel = 100
    }
    @{
        Name = 'Speeding fine $15'
        Money = -15
    }
    @{
        Name = 'Take a trip to Reading Railroad'
        Space = 5
    }
    @{
        Name = 'You have been elected Chairman of the Board. Pay each player $50.'
        Money = -50
        PlayerPayout = $true
    }
    @{
        Name = 'Your building loan matures. Collect $150!'
        Money = 150
    }
    @{
        Name = 'You have won a crossword competition! Collect $100!'
        Money = 100
    }
)
$ChanceCards = $ChanceCards | Get-Random -Count $ChanceCards.Count
$ChanceIndex = 0

# Community Chest Cards
$CommunityChestCards = @(
    @{
        Name = 'Advance to Go'
        Space = 0
    }
    @{
        Name = 'Bank error in your favor. Collect $200!'
        Money = 200
    }
    @{
        Name = 'Doctor fee. Pay $50.'
        Money = -50
    }
    @{
        Name = 'From sale of stock you get $50!'
        Money = 50
    }
    @{
        Name = 'Get Out of Jail Free'
    }
    @{
        Name = 'Go to Jail'
        Space = 10
        InJail = $true
    }
    @{
        Name = 'Holiday fund matures. Receive $100!'
        Money = 100
    }
    @{
        Name = 'Income tax refund. Collect $20!'
        Money = 20
    }
    @{
        Name = 'It is your birthday! Collect $10 from every player!'
        Money = 10
        PlayerPayout = $true
    }
    @{
        Name = 'Life insurance matures. Collect $100!'
        Money = 100
    }
    @{
        Name = 'Pay hospital fees of $100'
        Money = -100
    }
    @{
        Name = 'Pay school fees of $50'
        Money = -50
    }
    @{
        Name = 'Receive $25 consultancy fee!'
        Money = 25
    }
    @{
        Name = 'You are assessed for street repair. $40 per house. $115 per hotel.'
        MoneyPerHouse = 40
        MoneyPerHotel = 115
    }
    @{
        Name = 'You have won second prize in a beauty contest. Collect $10!'
        Money = 10
    }
    @{
        Name = 'You inherit $100!'
        Money = 100
    }
)
$CommunityChestCards = $CommunityChestCards | Get-Random -Count $CommunityChestCards.Count
$CommunityChestIndex = 0

# Create objects for spaces
$Spaces = @(
    @{
        Name = 'Go'
        Space = 0
        Owner = 'Go'
        X = 753
        Y = 753
    }
    @{
        Name = 'Mediterranean Avenue'
        Number = 1
        Cost = 60
        Mortgaged = $false
        Mortgage = 30
        UnmortgageCost = 33
        NumberInSet = 2
        Owner = 'Bank'
        Rent = 2
        Type = 'Property'
        SetColor = 'Purple'
        BuildingCost = 50
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(10,30,90,160,250)
        X = 647
        Y = 753
    }
    @{
        Name = 'Community Chest 1'
        Number = 2
        Owner = 'Gamble'
        Card = $CommunityChestCards
        X = 581
        Y = 753
    }
    @{
        Name = 'Baltic Avenue'
        Number = 3
        Cost = 60
        Mortgaged = $false
        Mortgage = 30
        UnmortgageCost = 33
        NumberInSet = 2
        Owner = 'Bank'
        Rent = 4
        Type = 'Property'
        SetColor = 'Purple'
        BuildingCost = 50
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(20,60,180,320,450)
        X = 516
        Y = 753
    }
    @{
        Name = 'Income Tax'
        Number = 4
        Owner = 'IRS'
        Rent = 200
        X = 450
        Y = 753
    }
    @{
        Name = 'Reading Railroad'
        Number = 5
        Cost = 200
        Mortgaged = $false
        Mortgage = 100
        UnmortgageCost = 110
        NumberInSet = 4
        SetColor = 'Black'
        Owner = 'Bank'
        Rent = 25
        Type = 'Railroad'
        X = 385
        Y = 753
    }
    @{
        Name = 'Oriental Avenue'
        Number = 6
        Cost = 100
        Mortgaged = $false
        Mortgage = 50
        UnmortgageCost = 55
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 6
        Type = 'Property'
        SetColor = 'Cyan'
        BuildingCost = 50
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(30,90,270,400,550)
        X = 319
        Y = 753
    }
    @{
        Name = 'Chance 1'
        Number = 7
        Owner = 'Gamble'
        Card = $ChanceCards
        X = 254
        Y = 753
    }
    @{
        Name = 'Vermont Avenue'
        Number = 8
        Cost = 100
        Mortgaged = $false
        Mortgage = 50
        UnmortgageCost = 55
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 6
        Type = 'Property'
        SetColor = 'Cyan'
        BuildingCost = 50
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(30,90,270,400,550)
        X = 188
        Y = 753
    }
    @{
        Name = 'Connecticut Avenue'
        Number = 9
        Cost = 120
        Mortgaged = $false
        Mortgage = 60
        UnmortgageCost = 66
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 8
        Type = 'Property'
        SetColor = 'Cyan'
        BuildingCost = 50
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(40,100,300,450,600)
        X = 123
        Y = 753
    }
    @{
        Name = 'Jail (Just Visiting)'
        Number = 10
        Owner = 'Null'
        X = 17
        Y = 753
    }
    @{
        Name = 'St. Charles Place'
        Number = 11
        Cost = 140
        Mortgaged = $false
        Mortgage = 70
        UnmortgageCost = 77
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 10
        Type = 'Property'
        SetColor = 'Magenta'
        BuildingCost = 100
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(50,150,450,625,750)
        X = 17
        Y = 647
    }
    @{
        Name = 'Electric Company'
        Number = 12
        Cost = 150
        Mortgaged = $false
        Mortgage = 75
        UnmortgageCost = 83
        NumberInSet = 2
        SetColor = 'White'
        Owner = 'Bank'
        Rent = $null
        Type = 'Utility'
        X = 17
        Y = 582
    }
    @{
        Name = 'States Avenue'
        Number = 13
        Cost = 140
        Mortgaged = $false
        Mortgage = 70
        UnmortgageCost = 77
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 10
        Type = 'Property'
        SetColor = 'Magenta'
        BuildingCost = 100
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(50,150,450,625,750)
        X = 17
        Y = 516
    }
    @{
        Name = 'Virginia Avenue'
        Number = 14
        Cost = 160
        Mortgaged = $false
        Mortgage = 80
        UnmortgageCost = 88
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 12
        Type = 'Property'
        SetColor = 'Magenta'
        BuildingCost = 100
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(60,180,500,700,900)
        X = 17
        Y = 451
    }
    @{
        Name = 'Pennsylvania Railroad'
        Number = 15
        Cost = 200
        Mortgaged = $false
        Mortgage = 100
        UnmortgageCost = 110
        NumberInSet = 4
        SetColor = 'Black'
        Owner = 'Bank'
        Rent = 25
        Type = 'Railroad'
        X = 17
        Y = 385
    }
    @{
        Name = 'St. James Place'
        Number = 16
        Cost = 180
        Mortgage = 90
        Mortgaged = $false
        UnmortgageCost = 99
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 14
        Type = 'Property'
        SetColor = 'Orange'
        BuildingCost = 100
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(70,200,550,750,950)
        X = 17
        Y = 319
    }
    @{
        Name = 'Community Chest 2'
        Number = 17
        Owner = 'Gamble'
        Card = $CommunityChestCards
        X = 17
        Y = 254
    }
    @{
        Name = 'Tennessee Avenue'
        Number = 18
        Cost = 180
        Mortgage = 90
        UnmortgageCost = 99
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 14
        Type = 'Property'
        SetColor = 'Orange'
        BuildingCost = 100
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(70,200,550,750,950)
        X = 17
        Y = 188
    }
    @{
        Name = 'New York Avenue'
        Number = 19
        Cost = 200
        Mortgaged = $false
        Mortgage = 100
        UnmortgageCost = 110
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 16
        Type = 'Property'
        SetColor = 'Orange'
        BuildingCost = 100
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(80,220,600,800,1000)
        X = 17
        Y = 123
    }
    @{
        Name = 'Free Parking'
        Number = 20
        Owner = 'Null'
        X = 17
        Y = 17
    }
    @{
        Name = 'Kentucky Avenue'
        Number = 21
        Cost = 220
        Mortgaged = $false
        Mortgage = 110
        UnmortgageCost = 121
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 18
        Type = 'Property'
        SetColor = 'Red'
        BuildingCost = 150
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(90,250,700,875,1050)
        X = 123
        Y = 17
    }
    @{
        Name = 'Chance 2'
        Number = 22
        Owner = 'Gamble'
        Card = $ChanceCards
        X = 188
        Y = 17
    }
    @{
        Name = 'Indiana Avenue'
        Number = 23
        Cost = 220
        Mortgaged = $false
        Mortgage = 110
        UnmortgageCost = 121
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 18
        Type = 'Property'
        SetColor = 'Red'
        BuildingCost = 150
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(90,250,700,875,1050)
        X = 254
        Y = 17
    }
    @{
        Name = 'Illinois Avenue'
        Number = 24
        Cost = 240
        Mortgaged = $false
        Mortgage = 120
        UnmortgageCost = 132
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 20
        Type = 'Property'
        SetColor = 'Red'
        BuildingCost = 150
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(100,300,750,925,1100)
        X = 319
        Y = 17
    }
    @{
        Name = 'B&O Railroad'
        Number = 25
        Cost = 200
        Mortgaged = $false
        Mortgage = 100
        UnmortgageCost = 110
        NumberInSet = 4
        SetColor = 'Black'
        Owner = 'Bank'
        Rent = 25
        Type = 'Railroad'
        X = 385
        Y = 17
    }
    @{
        Name = 'Atlantic Avenue'
        Number = 26
        Cost = 260
        Mortgaged = $false
        Mortgage = 130
        UnmortgageCost = 143
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 22
        Type = 'Property'
        SetColor = 'Yellow'
        BuildingCost = 150
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(110,330,800,975,1150)
        X = 450
        Y = 17
    }
    @{
        Name = 'Ventnor Avenue'
        Number = 27
        Cost = 260
        Mortgaged = $false
        Mortgage = 130
        UnmortgageCost = 143
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 22
        Type = 'Property'
        SetColor = 'Yellow'
        BuildingCost = 150
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(110,330,800,975,1150)
        X = 516
        Y = 17
    }
    @{
        Name = 'Water Works'
        Number = 28
        Cost = 150
        Mortgaged = $false
        Mortgage = 75
        UnmortgageCost = 83
        NumberInSet = 2
        SetColor = 'White'
        Owner = 'Bank'
        Rent = $null
        Type = 'Utility'
        X = 582
        Y = 17
    }
    @{
        Name = 'Marvin Gardens'
        Number = 29
        Cost = 280
        Mortgaged = $false
        Mortgage = 140
        UnmortgageCost = 154
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 24
        Type = 'Property'
        SetColor = 'Yellow'
        BuildingCost = 150
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(120,360,850,1025,1200)
        X = 647
        Y = 17
    }
    @{
        Name = 'Go To Jail'
        Number = 30
        Owner = 'Jail'
        X = 753
        Y = 17
    }
    @{
        Name = 'Pacific Avenue'
        Number = 31
        Cost = 300
        Mortgaged = $false
        Mortgage = 150
        UnmortgageCost = 165
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 26
        Type = 'Property'
        SetColor = 'Green'
        BuildingCost = 200
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(130,390,900,1100,1275)
        X = 753
        Y = 123
    }
    @{
        Name = 'North Carolina Avenue'
        Number = 32
        Cost = 300
        Mortgaged = $false
        Mortgage = 150
        UnmortgageCost = 165
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 26
        Type = 'Property'
        SetColor = 'Green'
        BuildingCost = 200
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(130,390,900,1100,1275)
        X = 753
        Y = 188
    }
    @{
        Name = 'Community Chest 3'
        Number = 33
        Owner = 'Gamble'
        Card = $CommunityChestCards
        X = 753
        Y = 253
    }
    @{
        Name = 'Pennsylvania Avenue'
        Number = 34
        Cost = 320
        Mortgaged = $false
        Mortgage = 160
        UnmortgageCost = 176
        NumberInSet = 3
        Owner = 'Bank'
        Rent = 28
        Type = 'Property'
        SetColor = 'Green'
        BuildingCost = 200
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(150,450,1000,1200,1400)
        X = 753
        Y = 319
    }
    @{
        Name = 'Short Line'
        Number = 35
        Cost = 200
        Mortgaged = $false
        Mortgage = 100
        UnmortgageCost = 110
        NumberInSet = 4
        SetColor = 'Black'
        Owner = 'Bank'
        Rent = 25
        Type = 'Railroad'
        X = 753
        Y = 385
    }
    @{
        Name = 'Chance 3'
        Number = 36
        Owner = 'Gamble'
        Card = $ChanceCards
        X = 753
        Y = 451
    }
    @{
        Name = 'Park Place'
        Number = 37
        Cost = 350
        Mortgaged = $false
        Mortgage = 175
        UnmortgageCost = 193
        NumberInSet = 2
        Owner = 'Bank'
        Rent = 35
        Type = 'Property'
        SetColor = 'Blue'
        BuildingCost = 200
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(175,500,1100,1300,1500)
        X = 753
        Y = 516
    }
    @{
        Name = 'Luxury Tax'
        Number = 38
        Rent = 100
        Owner = 'IRS'
        X = 753
        Y = 582
    }
    @{
        Name = 'Boardwalk'
        Number = 39
        Cost = 400
        Mortgaged = $false
        Mortgage = 200
        UnmortgageCost = 220
        NumberInSet = 2
        Owner = 'Bank'
        Rent = 50
        Type = 'Property'
        SetColor = 'Blue'
        BuildingCost = 200
        HouseCount = 0
        HotelCount = 0
        BuildingRent = @(200,600,1400,1700,2000)
        X = 753
        Y = 647
    }
)

# Create game board
$GameBoardIcon = "$PSScriptRoot\MonopolyBoard.ico"
$GameBoardImage = "$PSScriptRoot\MonopolyBoard.jpg"
$ScreenWidth = [System.Windows.SystemParameters]::PrimaryScreenWidth
$ScreenLeft = $ScreenWidth - 900
$Xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Icon="$GameBoardIcon"
        Title='Monopoly Board'
        Left="$ScreenLeft"
        Top='100'
        WindowStyle='None'
        ResizeMode='NoResize'
        SizeToContent='WidthAndHeight'
        Topmost='True'>
    <Window.TaskbarItemInfo>
        <TaskbarItemInfo/>
    </Window.TaskbarItemInfo>
    <Border BorderBrush='Gray' BorderThickness='5'>
        <Grid>
            <Image Width='800' Height='800' Source='$GameBoardImage' Stretch='None'/>
            <Canvas Name='BoardCanvas'>
                <Canvas Name='HousesAndHotelsCanvas'/>
                <Canvas Name='PropertyOwnerCanvas'/>
            </Canvas>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition/>
                </Grid.ColumnDefinitions>
                <Image x:Name="DieImage1"
                       Width="50"
                       Height="50"
                       HorizontalAlignment="Center"
                       VerticalAlignment="Center"/>
                <Image x:Name="DieImage2"
                       Width="50"
                       Height="50"
                       HorizontalAlignment="Center"
                       VerticalAlignment="Center"
                       Grid.Column="1" />
            </Grid>
        </Grid>
    </Border>
</Window>
"@

# Open game board
$Window = [Windows.Markup.XamlReader]::Parse($Xaml)
$null = Register-ObjectEvent -InputObject $Window -EventName 'Closed' -Action { $Window.Dispatcher.InvokeShutdown() }
$OS = Get-CimInstance Win32_OperatingSystem
if ($OS.BuildNumber -gt 22000) {
    $Window.TaskbarItemInfo.Overlay = $GameBoardIcon
}
$Window.Show()
$Window.Add_MouseLeftButtonDown({
    $Window.DragMove()
})

# Exit the game function
function ExitGame {
    $Window.TaskbarItemInfo.Overlay = $null
    $Window.Dispatcher.Invoke([Action]{$Window.Close()})
}

# Dice controls
$DieImage1 = $Window.FindName("DieImage1")
$DieImage2 = $Window.FindName("DieImage2")

# Create player pieces
$Spacing = 0
$PieceSize = 30
#$BoardCanvas = $Window.FindName('BoardCanvas')
$PlayerGrids = @{}

# Update player pieces function
function UpdatePlayerPieces {
    $BoardCanvas = $Window.FindName('BoardCanvas')
    $AvailableIndexes = 0..7 | Get-Random -Count 8
    
    foreach ($Player in $Players) {
        $PlayerGrid = $PlayerGrids[$Player.Name]
        if ($PlayerGrid -eq $null) {
            $RandomIndex = $AvailableIndexes[0]
            $AvailableIndexes = $AvailableIndexes | Where-Object { $_ -ne $RandomIndex }
            $PlayerGrid = New-Object System.Windows.Controls.Grid
            $PlayerGrids[$Player.Name] = $PlayerGrid

            $PlayerShape = $null
            switch ($RandomIndex) {
                0 { # Circle
                    $PlayerShape = New-Object System.Windows.Shapes.Ellipse
                }
                1 { # Square
                    $PlayerShape = New-Object System.Windows.Shapes.Rectangle
                }
                2 { # Triangle
                    $PlayerShape = New-Object System.Windows.Shapes.Polygon
                    $PlayerShape.Points = New-Object System.Windows.Media.PointCollection
                    $PlayerShape.Points.Add((New-Object System.Windows.Point -ArgumentList 0, $PieceSize))
                    $PlayerShape.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize / 2), 0))
                    $PlayerShape.Points.Add((New-Object System.Windows.Point -ArgumentList $PieceSize, $PieceSize))
                }
                3 { # Rhombus
                    $PlayerShape = New-Object System.Windows.Shapes.Path
                    $PathGeometry = New-Object System.Windows.Media.PathGeometry
                    $PathFigure = New-Object System.Windows.Media.PathFigure
                    $PathFigure.StartPoint = New-Object System.Windows.Point -ArgumentList ($PieceSize / 2), 0
                    $PathFigure.IsClosed = $true
                    $PathGeometry.Figures.Add($PathFigure)

                    $RhombusSegment = New-Object System.Windows.Media.PolyLineSegment
                    $RhombusSegment.Points = New-Object System.Windows.Media.PointCollection
                    $RhombusSegment.Points.Add((New-Object System.Windows.Point -ArgumentList $PieceSize, ($PieceSize / 2)))
                    $RhombusSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize / 2), $PieceSize))
                    $RhombusSegment.Points.Add((New-Object System.Windows.Point -ArgumentList 0, ($PieceSize / 2)))

                    $PathFigure.Segments.Add($RhombusSegment)
                    $PlayerShape.Data = $PathGeometry
                }
                4 { # Star
                    $PlayerShape = New-Object System.Windows.Shapes.Path
                    $PathGeometry = New-Object System.Windows.Media.PathGeometry
                    $PathFigure = New-Object System.Windows.Media.PathFigure
                    $PathFigure.StartPoint = New-Object System.Windows.Point -ArgumentList ($PieceSize / 2), 0
                    $PathFigure.IsClosed = $true
                    $PathGeometry.Figures.Add($PathFigure)

                    $StarSegment = New-Object System.Windows.Media.PolyLineSegment
                    $StarSegment.Points = New-Object System.Windows.Media.PointCollection
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.62), ($PieceSize * 0.38)))
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.98), ($PieceSize * 0.38)))
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.67), ($PieceSize * 0.61)))
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.78), ($PieceSize * 0.97)))
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.5), ($PieceSize * 0.73)))
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.22), ($PieceSize * 0.97)))
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.33), ($PieceSize * 0.61)))
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList 0.02, ($PieceSize * 0.38)))
                    $StarSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.38), ($PieceSize * 0.38)))

                    $PathFigure.Segments.Add($StarSegment)
                    $PlayerShape.Data = $PathGeometry
                }
                5 { # Pentagon
                    $PlayerShape = New-Object System.Windows.Shapes.Path
                    $PathGeometry = New-Object System.Windows.Media.PathGeometry
                    $PathFigure = New-Object System.Windows.Media.PathFigure
                    $PathFigure.StartPoint = New-Object System.Windows.Point -ArgumentList ($PieceSize / 2), 0
                    $PathFigure.IsClosed = $true
                    $PathGeometry.Figures.Add($PathFigure)

                    $PentagonSegment = New-Object System.Windows.Media.PolyLineSegment
                    $PentagonSegment.Points = New-Object System.Windows.Media.PointCollection
                    $PentagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList $PieceSize, ($PieceSize * 0.381)))
                    $PentagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.809), $PieceSize))
                    $PentagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.191), $PieceSize))
                    $PentagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList 0, ($PieceSize * 0.381)))

                    $PathFigure.Segments.Add($PentagonSegment)
                    $PlayerShape.Data = $PathGeometry
                }
                6 { # Hexagon
                    $PlayerShape = New-Object System.Windows.Shapes.Path
                    $PathGeometry = New-Object System.Windows.Media.PathGeometry
                    $PathFigure = New-Object System.Windows.Media.PathFigure
                    $PathFigure.StartPoint = New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.25), 0
                    $PathFigure.IsClosed = $true
                    $PathGeometry.Figures.Add($PathFigure)

                    $HexagonSegment = New-Object System.Windows.Media.PolyLineSegment
                    $HexagonSegment.Points = New-Object System.Windows.Media.PointCollection
                    $HexagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.75), 0))
                    $HexagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList $PieceSize, ($PieceSize * 0.5)))
                    $HexagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.75), $PieceSize))
                    $HexagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.25), $PieceSize))
                    $HexagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList 0, ($PieceSize * 0.5)))

                    $PathFigure.Segments.Add($HexagonSegment)
                    $PlayerShape.Data = $PathGeometry
                }
                7 { # Octagon
                    $PlayerShape = New-Object System.Windows.Shapes.Path
                    $PathGeometry = New-Object System.Windows.Media.PathGeometry
                    $PathFigure = New-Object System.Windows.Media.PathFigure
                    $PathFigure.StartPoint = New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.2929), 0
                    $PathFigure.IsClosed = $true
                    $PathGeometry.Figures.Add($PathFigure)

                    $OctagonSegment = New-Object System.Windows.Media.PolyLineSegment
                    $OctagonSegment.Points = New-Object System.Windows.Media.PointCollection
                    $OctagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.7071), 0))
                    $OctagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList $PieceSize, ($PieceSize * 0.2929)))
                    $OctagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList $PieceSize, ($PieceSize * 0.7071)))
                    $OctagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.7071), $PieceSize))
                    $OctagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList ($PieceSize * 0.2929), $PieceSize))
                    $OctagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList 0, ($PieceSize * 0.7071)))
                    $OctagonSegment.Points.Add((New-Object System.Windows.Point -ArgumentList 0, ($PieceSize * 0.2929)))

                    $PathFigure.Segments.Add($OctagonSegment)
                    $PlayerShape.Data = $PathGeometry
                }
            }
            
            $PlayerShape.Width = $PieceSize
            $PlayerShape.Height = $PieceSize
            $PlayerShape.Fill = $Player.Color
            $PlayerShape.Stroke = 'Black'

            $PlayerTextBlock = New-Object System.Windows.Controls.TextBlock
            $PlayerTextBlock.Text = $Player.Name.Substring(0, 1)
            $PlayerTextBlock.VerticalAlignment = 'Center'
            $PlayerTextBlock.HorizontalAlignment = 'Center'
            $PlayerTextBlock.FontWeight = 'Bold'
            $PlayerTextBlock.Foreground = 'Black'

            $null = $PlayerGrid.Children.Add($PlayerShape)
            $null = $PlayerGrid.Children.Add($PlayerTextBlock)
            $null = $BoardCanvas.Children.Add($PlayerGrid)
        }

        $PlayerCurrentSpace = $Player.CurrentSpace
        $PlayersOnSpace = @()
        $PlayersInJail = @()
        $PlayersVisiting = @()

        foreach ($OtherPlayer in $Players) {
            if ($OtherPlayer.CurrentSpace -eq $PlayerCurrentSpace) {
                $PlayersOnSpace += $OtherPlayer
                if ($PlayerCurrentSpace -eq 10) {
                    if ($OtherPlayer.InJail) {
                        $PlayersInJail += $OtherPlayer
                    } else {
                        $PlayersVisiting += $OtherPlayer
                    }
                }
            }
        }

        if ($PlayerGrid -ne $null) {
            if ($PlayerCurrentSpace -eq 10) {
                $PlayerIndex = [array]::IndexOf($PlayersInJail, $Player)
                $TotalPlayers = $PlayersInJail.Count
                if (-not $Player.InJail) {
                    $PlayerIndex = [array]::IndexOf($PlayersVisiting, $Player)
                    $TotalPlayers = $PlayersVisiting.Count
                }
            } else {
                $PlayerIndex = [array]::IndexOf($PlayersOnSpace, $Player)
                $TotalPlayers = $PlayersOnSpace.Count
            }

            $SpacingCalc = 22 / $TotalPlayers
            $AdjustedSpacing = $SpacingCalc * ($TotalPlayers - 1)

            if ($Player.InJail) {
                $PlayerGridX = 53 - $AdjustedSpacing + ($PlayerIndex * $SpacingCalc * 2)
                $PlayerGridY = 718 - $AdjustedSpacing + ($PlayerIndex * $SpacingCalc * 2)
            } else {
                $PlayerGridX = $Spaces[$PlayerCurrentSpace].X - $AdjustedSpacing + ($PlayerIndex * $SpacingCalc * 2)
                $PlayerGridY = $Spaces[$PlayerCurrentSpace].Y - $AdjustedSpacing + ($PlayerIndex * $SpacingCalc * 2)
            }
            
            [Windows.Controls.Canvas]::SetLeft($PlayerGrid, $PlayerGridX)
            [Windows.Controls.Canvas]::SetTop($PlayerGrid, $PlayerGridY)
            
            $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        }
    }
}

# Show property owner function
function UpdatePropertyOwner {
    $PropertyOwnerCanvas = $Window.FindName('PropertyOwnerCanvas')
    $PropertyOwnerCanvas.Children.Clear()

    foreach ($Space in $Spaces) {
        if ($Players.Name -contains $Space.Owner) {
            $OwnerMarkerColor = ($Players | Where { $_.Name -eq $Space.Owner}).Color
                        
            switch ($Space.Y)
            {
                753 {
                    $Side = "Bottom"
                    $OwnerMarkerX = $Space.X - 18
                    $OwnerMarkerY = $Space.Y - 64
                    $OwnerMarkerWidth = 66
                    $OwnerMarkerHeight = 5
                }
                17  {
                    $Side = "Top"
                    $OwnerMarkerX = $Space.X - 18
                    $OwnerMarkerY = $Space.Y + 74 + ($PieceSize / 2)
                    $OwnerMarkerWidth = 66
                    $OwnerMarkerHeight = 5
                }
            }
            
            switch ($Space.X)
            {
                17 {
                    $Side = "Left"
                    $OwnerMarkerX = $Space.X + 74 + ($PieceSize / 2)
                    $OwnerMarkerY = $Space.Y - 18
                    $OwnerMarkerWidth = 5
                    $OwnerMarkerHeight = 66
                }
                753  {
                    $Side = "Right"
                    $OwnerMarkerX = $Space.X - 64
                    $OwnerMarkerY = $Space.Y - 18
                    $OwnerMarkerWidth = 5
                    $OwnerMarkerHeight = 66
                }
            }

            $OwnerMarker = New-Object Windows.Shapes.Rectangle
            $OwnerMarker.Width = $OwnerMarkerWidth
            $OwnerMarker.Height = $OwnerMarkerHeight
            $OwnerMarker.Fill = $OwnerMarkerColor
            $OwnerMarker.Stroke = 'Black'
            [Windows.Controls.Canvas]::SetLeft($OwnerMarker, $OwnerMarkerX)
            [Windows.Controls.Canvas]::SetTop($OwnerMarker, $OwnerMarkerY)
            $null = $PropertyOwnerCanvas.Children.Add($OwnerMarker)

            $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        }
    }
}

# Update properties function
function UpdateProperties {
    $HousesAndHotelsCanvas = $Window.FindName('HousesAndHotelsCanvas')
    $HousesAndHotelsCanvas.Children.Clear()
    $HouseWidth = 15
    $HouseHeight = 15

    foreach ($Space in $Spaces) {
        if ($Space.Y -eq 753) {$Side = "Bottom"}
        if ($Space.Y -eq 17) {$Side = "Top"}
        if ($Space.X -eq 17) {$Side = "Left"}
        if ($Space.X -eq 753) {$Side = "Right"}

        if ($Space.Type -eq 'Property') {
            $HouseCount = $Space.HouseCount
            $HotelCount = $Space.HotelCount
            
            switch ($Side) {
                Bottom {
                    $HouseX = $Space.X - 16
                    $HouseY = $Space.Y - 54
                    $HotelX = $HouseX + 16
                    $HotelY = $HouseY
                    $HotelWidth = 30
                    $HotelHeight = 15
                }
                Top  {
                    $HouseX = $Space.X + 16 + ($PieceSize / 2)
                    $HouseY = $Space.Y + 54 + ($PieceSize / 2)
                    $HotelX = $HouseX - 30
                    $HotelY = $HouseY
                    $HotelWidth = 30
                    $HotelHeight = 15
                }
                Left {
                    $HouseX = $Space.X + 54 + ($PieceSize / 2)
                    $HouseY = $Space.Y - 16
                    $HotelX = $HouseX
                    $HotelY = $HouseY + 16
                    $HotelWidth = 15
                    $HotelHeight = 30
                }
                Right  {
                    $HouseX = $Space.X - 54
                    $HouseY = $Space.Y + 16 + ($PieceSize / 2)
                    $HotelX = $HouseX
                    $HotelY = $HouseY - 30
                    $HotelWidth = 15
                    $HotelHeight = 30
                }
            }

            for ($i = 0; $i -lt $HouseCount; $i++) {
                $HouseRectangle = New-Object Windows.Shapes.Rectangle
                $HouseRectangle.Width = $HouseWidth
                $HouseRectangle.Height = $HouseHeight
                $HouseRectangle.Fill = 'Green'
                $HouseRectangle.Stroke = 'Gray'
                switch ($Side) {
                    Bottom {
                        [Windows.Controls.Canvas]::SetLeft($HouseRectangle, $HouseX + $i * ($HouseWidth + 1))
                        [Windows.Controls.Canvas]::SetTop($HouseRectangle, $HouseY)
                    }
                    Top {
                        [Windows.Controls.Canvas]::SetLeft($HouseRectangle, $HouseX - $i * ($HouseWidth + 1))
                        [Windows.Controls.Canvas]::SetTop($HouseRectangle, $HouseY)
                    }
                    Left {
                        [Windows.Controls.Canvas]::SetLeft($HouseRectangle, $HouseX)
                        [Windows.Controls.Canvas]::SetTop($HouseRectangle, $HouseY + $i * ($HouseWidth + 1))
                    }
                    Right {
                        [Windows.Controls.Canvas]::SetLeft($HouseRectangle, $HouseX)
                        [Windows.Controls.Canvas]::SetTop($HouseRectangle, $HouseY - $i * ($HouseWidth + 1))
                    }
                }
                $null = $HousesAndHotelsCanvas.Children.Add($HouseRectangle)
            }

            if ($HotelCount -gt 0) {
                $HotelRectangle = New-Object Windows.Shapes.Rectangle
                $HotelRectangle.Width = $HotelWidth
                $HotelRectangle.Height = $HotelHeight
                $HotelRectangle.Fill = 'Red'
                $HotelRectangle.Stroke = 'Gray'
                [Windows.Controls.Canvas]::SetLeft($HotelRectangle, $HotelX)
                [Windows.Controls.Canvas]::SetTop($HotelRectangle, $HotelY)
                $null = $HousesAndHotelsCanvas.Children.Add($HotelRectangle)

                $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
            }
        }

        # Add "MORTGAGED" label if property is mortgaged
        if ($Space.Mortgaged) {
            $MortgagedLabel = New-Object Windows.Controls.TextBlock
            $MortgagedLabel.Text = "MORTGAGED"
            $MortgagedLabel.FontSize = 9
            #$MortgagedLabel.Height = $HouseHeight
            $MortgagedLabel.Foreground = 'Black'
            $MortgagedLabel.FontWeight = 'Bold'
            $MortgagedLabel.VerticalAlignment = 'Center'
            $MortgagedLabel.HorizontalAlignment = 'Center'

            switch ($Side) {
                Bottom {
                    [Windows.Controls.Canvas]::SetLeft($MortgagedLabel, $Space.X - 13)
                    [Windows.Controls.Canvas]::SetTop($MortgagedLabel, $HouseY + 1)
                }
                Top {
                    [Windows.Controls.Canvas]::SetLeft($MortgagedLabel, $Space.X - 13)
                    [Windows.Controls.Canvas]::SetTop($MortgagedLabel, $HouseY + 1)
                    $MortgagedLabel.LayoutTransform = New-Object Windows.Media.RotateTransform(180)
                }
                Left {
                    [Windows.Controls.Canvas]::SetLeft($MortgagedLabel, $HouseX + 1)
                    [Windows.Controls.Canvas]::SetTop($MortgagedLabel, $Space.Y - 13)
                    $MortgagedLabel.LayoutTransform = New-Object Windows.Media.RotateTransform(90)
                }
                Right {
                    [Windows.Controls.Canvas]::SetLeft($MortgagedLabel, $HouseX + 1)
                    [Windows.Controls.Canvas]::SetTop($MortgagedLabel, $Space.Y - 13)
                    $MortgagedLabel.LayoutTransform = New-Object Windows.Media.RotateTransform(270)
                }
            }
            $null = $HousesAndHotelsCanvas.Children.Add($MortgagedLabel)

            $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        }
    }
}

# Main game loop
Write-Host `n
Write-Host "Here we go!"`n
UpdatePlayerPieces
$PlayerIndex = 0
while (($Players | Measure-Object).Count -gt 1) {
    $Player = $Players[$PlayerIndex]
    do {
        # Start of turn
        Write-Host "$($Player.Name)'s turn! `(`$$($Player.Money)`) " -NoNewLine -ForegroundColor $Player.Color
        if ($Player.IsHuman) {
            Write-Host "Press SPACEBAR to roll the dice"`n
            do {
                while ([System.Console]::KeyAvailable) {
                    [System.Console]::ReadKey($true) | Out-Null
                }
                $KeyInfo = [System.Console]::ReadKey($true)
                $KeyChar = $KeyInfo.KeyChar
            } while ($KeyChar -ne ' ')
        } else {
            Write-Host `n
        }
        $DiceRoll = DiceRoll
        $MoveCount = $DiceRoll[0]
        $Double = $DiceRoll[1]
        Write-Host "$($Player.Name) rolled $MoveCount"`n -ForegroundColor $Player.Color
        if (!$Testing) {Start-Sleep -Milliseconds $Pause}

        if ($Player.InJail) {
            if ($Double) {
                Write-Host "$($Player.Name) rolled a double! You get out of jail!"`n -ForegroundColor $Player.Color
                $Player.InJail = $false
                $Tries = 0
                $Double = $false
                if (!$Testing) {Start-Sleep -Milliseconds $Pause}
            } else {
                $Tries++
                if ($Tries -lt 3) {
                    Write-Host "$($Player.Name) did not roll a double. $(3 - $Tries) more chances to roll a double, or pay `$50 to get out."`n -ForegroundColor $Player.Color
                    if ($Player.GetOutOfJailFree -gt 0) {
                        Write-Host "$($Player.Name) used a Get Out Of Jail Free card!"`n -ForegroundColor $Player.Color
                        $Player.InJail = $false
                        $Tries = 0
                        $Player.GetOutOfJailFree--
                        if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    }
                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                } else {
                    Write-Host "$($Player.Name) has failed to roll a double in 3 tries. Pay `$50 to get out."`n -ForegroundColor $Player.Color
                    $Player.InJail = $false
                    $Tries = 0
                    $Player.Money -= 50
                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    $BankruptcyCheck = BankruptcyCheck
                    $Players = $BankruptcyCheck.RemainingPlayers
                    if ($BankruptcyCheck.IsBankrupt) {
                        Write-Host "$($Player.Name) has gone bankrupt!"`n -ForegroundColor $Player.Color
                        $PlayerIndex = ($PlayerIndex += 1) % ($Players | Measure-Object).Count
                        continue
                    }
                }
            }
        }
        if (!$Player.InJail) {
            if ($Double) {
                if ($Player.DoubleCount -lt 2) {
                    $Player.DoubleCount++
                    Write-Host "$($Player.Name) rolled a double! $($Player.Name) will get to roll again!"`n -ForegroundColor $Player.Color
                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                } else {
                    Write-Host "$($Player.Name) rolled 3 doubles in a row! Go straight to Jail you little cheater!"`n -ForegroundColor $Player.Color
                    $Player.CurrentSpace = 10
                    $Player.InJail = $true
                    $Player.DoubleCount = 0
                    $Double = $false
                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    break
                }
            } else {
                $Player.DoubleCount = 0
            }

            # Move player according to dice roll
            $PreviousSpace = $Player.CurrentSpace
            for ($i = 0; $i -le $MoveCount; $i++) {
                $Player.CurrentSpace = ($PreviousSpace + $i) % $Spaces.Count
                $CurrentSpace = $Spaces[$Player.CurrentSpace]
                UpdatePlayerPieces
                if (!$Testing) {Start-Sleep -Milliseconds 500}
            }
            $SpaceColor = GetPropertyColor $CurrentSpace.SetColor
            if ($Player.InJail) {
                Write-Host "$($Player.Name) went to Jail!"`n -ForegroundColor $Player.Color
            } else {
                Write-Host "$($Player.Name) landed on " -NoNewline -ForegroundColor $Player.Color
                if ($CurrentSpace.SetColor -eq 'Orange' -and !$psISE) {
                    Write-Host $SpaceColor"$($CurrentSpace.Name)"`n$ResetColor
                } else {
                    Write-Host "$($CurrentSpace.Name)"`n -ForegroundColor $SpaceColor
                }
            }
            if (!$Testing) {Start-Sleep -Milliseconds $Pause}

            function SpaceActions { # Space actions
                # Did player pass go?
                if ($CurrentSpace.Name -ne 'Go' -and $CurrentSpace.Number -lt $PreviousSpace -and !$Player.InJail) {
                    Write-Host "$($Player.Name) passed Go! Collect `$200!"`n -ForegroundColor $Player.Color
                    $Player.Money += 200
                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                }

                switch ($CurrentSpace.Owner) {
                    Default {
                        # Rent logic - Utilies
                        if ($CurrentSpace.Type -eq 'Utility') {
                            if ($Spaces[12].Owner -eq $CurrentSpace.Owner -and $Spaces[28].Owner -eq $CurrentSpace.Owner) {
                                $Rent = ($MoveCount * 10)
                            } else {
                                $Rent = ($MoveCount * 4)
                            }
                        }

                        # Rent logic - Railroads
                        if ($CurrentSpace.Type -eq 'Railroad') {
                            $RailroadsOwned = (($Spaces | Where {$_.Owner -eq $CurrentSpace.Owner -and $_.Type -eq 'Railroad'}) | Measure-Object).Count
                            switch ($RailroadsOwned) {
                                1 {$Rent = 25}
                                2 {$Rent = 50}
                                3 {$Rent = 100}
                                4 {$Rent = 200}
                            }
                            if ($PayDouble) {
                                Write-Host "$($Player.Name) owes double rent!"`n -ForegroundColor $Player.Color
                                $Rent = $Rent * 2
                                $PayDouble = $false
                                if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                            }
                        }

                        # Rent logic - Properties
                        if ($CurrentSpace.Type -eq 'Property') {
                            $PropertyCount = (($Spaces | Where {$_.Owner -eq $CurrentSpace.Owner -and $_.SetColor -eq $CurrentSpace.SetColor}) | Measure-Object).Count
                            if ($PropertyCount -eq $CurrentSpace.NumberInSet) {
                                if ($CurrentSpace.HouseCount -gt 0 -or $CurrentSpace.HotelCount -gt 0) {
                                    if ($CurrentSpace.HouseCount -gt 0) {
                                        $Rent = $CurrentSpace.BuildingRent[($CurrentSpace.HouseCount - 1)]
                                    }
                                    if ($CurrentSpace.HotelCount -gt 0) {
                                        $Rent = $CurrentSpace.BuildingRent[4]
                                    }
                                } else {
                                    $Rent = ($CurrentSpace.Rent * 2)
                                }
                            } else {
                                $Rent = $CurrentSpace.Rent
                            }
                        }

                        Write-Host "This property is owned by " -NoNewline -ForegroundColor $Player.Color
                        Write-Host "$($CurrentSpace.Owner) " -NoNewline -ForegroundColor ($Players | Where {$_.Name -eq $CurrentSpace.Owner}).Color
                        if ($CurrentSpace.Mortgaged) {
                            Write-Host "but it is currenty mortgaged"`n -ForegroundColor $Player.Color
                        } else {
                        Write-Host "- pay them `$$($Rent)"`n -ForegroundColor $Player.Color
                            if ($Player.Money -lt $Rent) {
                                $Player.Money -= $Rent
                                $null = BankruptcyCheck
                                if ($Player.Money -lt 0) {
                                    $Rent = $Rent + $Player.Money # Adjust rent if player went bankrupt and can't pay whole amount
                                }
                            } else {
                                $Player.Money -= $Rent
                            }
                            ($Players | Where {$_.Name -eq $CurrentSpace.Owner}).Money += $Rent
                        }
                        if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    }

                    Bank {
                        if ($Player.Money -ge $CurrentSpace.Cost) {
                            if ($Player.IsHuman) {
                                Write-Host "Would you like to buy this property for `$$($CurrentSpace.Cost)? [Y/N]"
                                $Purchase = GetKeyPress
                            } else {
                                $Purchase = $true
                            }
                            if ($Purchase -eq $true) {
                                Write-Host "$($Player.Name) has purchased it for `$$($CurrentSpace.Cost)"`n -ForegroundColor $Player.Color
                                $CurrentSpace.Owner = $Player.Name
                                UpdatePropertyOwner
                                $Player.Money -= $CurrentSpace.Cost
                            }
                            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                        } else {
                            Write-Host "$($Player.Name) can't afford to buy this property"`n -ForegroundColor $Player.Color
                            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                        }
                    }

                    $Player.Name {
                        Write-Host "$($Player.Name) owns this property!"`n -ForegroundColor $Player.Color
                        if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    }

                    Go {
                        Write-Host 'Collect $200!'`n -ForegroundColor $Player.Color
                        $Player.Money += 200
                        if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    }

                    Gamble { # Chance and Community Chest cards
                        if ($CurrentSpace.Name -match 'Chance') {
                            $Card = $ChanceCards[$ChanceIndex]
                            $script:ChanceIndex = ($ChanceIndex + 1) % $ChanceCards.Count
                        }
                        if ($CurrentSpace.Name -match 'Community Chest') {
                            $Card = $CommunityChestCards[$CommunityChestIndex]
                            $script:CommunityChestIndex = ($CommunityChestIndex + 1) % $CommunityChestCards.Count
                        }                                
                        Write-Host Card: $Card.Name`n -ForegroundColor $Player.Color
                        if (!$Testing) {Start-Sleep -Milliseconds $Pause}

                        # Pay or earn money
                        if ($Card.Money -is [int]) {
                            if ($Card.PlayerPayout -is [int]) {
                                $Player.Money += $Card.Money * ($Players.Count - 1)
                                foreach ($OtherPlayer in $Players | Where {$_.Name -ne $Player.Name}) {
                                    $OtherPlayer.Money -= $Card.Money
                                }
                            } else {
                                $Player.Money += $Card.Money
                            }
                        }

                        # Move to space
                        if ($Card.Space -is [int]) {
                            $Player.CurrentSpace = $Card.Space
                            $CurrentSpace = $Spaces[$Player.CurrentSpace]
                            if ($Card.InJail) {
                                $Player.InJail = $true
                                $CurrentSpace.Name = "Jail"
                            }
                            $SpaceColor = GetPropertyColor $CurrentSpace.SetColor
                            UpdatePlayerPieces
                            Write-Host "$($Player.Name) moved to " -NoNewline -ForegroundColor $Player.Color
                            if ($CurrentSpace.SetColor -eq 'Orange') {
                                Write-Host $SpaceColor"$($CurrentSpace.Name)"`n$ResetColor
                            } else {
                                Write-Host "$($CurrentSpace.Name)"`n -ForegroundColor $SpaceColor
                            }
                            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                            SpaceActions
                        }

                        # Move relative spaces
                        if ($Card.SpaceMove -is [int]) {
                            $Player.CurrentSpace += $Card.SpaceMove
                            $CurrentSpace = $Spaces[$Player.CurrentSpace]
                            $SpaceColor = GetPropertyColor $CurrentSpace.SetColor
                            UpdatePlayerPieces
                            Write-Host "$($Player.Name) moved to " -NoNewline -ForegroundColor $Player.Color
                            if ($CurrentSpace.SetColor -eq 'Orange') {
                                Write-Host $SpaceColor"$($CurrentSpace.Name)"`n$ResetColor
                            } else {
                                Write-Host "$($CurrentSpace.Name)"`n -ForegroundColor $SpaceColor
                            }
                            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                            SpaceActions
                        }

                        # Pay for houses/hotels
                        if ($Card.MoneyPerHouse -is [int]) {
                            $HouseMaintenance = $Player.HousesOwned * $Card.MoneyPerHouse
                            $HotelMaintenance = $Player.HotelsOwned * $Card.MoneyPerHotel
                            $TotalMaintenance = $HouseMaintenance + $HotelMaintenance
                            if ($TotalMaintenance -gt 0) {
                                Write-Host "$($Player.Name) must pay `$$($TotalMaintenance)!"`n -ForegroundColor $Player.Color
                                $Player.Money -= $TotalMaintenance
                                $BankruptcyCheck = BankruptcyCheck
                                $Players = $BankruptcyCheck.RemainingPlayers
                                if ($BankruptcyCheck.IsBankrupt) {
                                    Write-Host "$($Player.Name) has gone bankrupt!"`n -ForegroundColor $Player.Color
                                    $PlayerIndex = ($PlayerIndex += 1) % ($Players | Measure-Object).Count
                                    continue
                                }
                            } else {
                                Write-Host "Luckily, $($Player.Name) doesn't own any houses or hotels!"`n -ForegroundColor $Player.Color
                            }
                            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                        }

                        # Get out of Jail free
                        if ($Card.Name -eq 'Get Out of Jail Free') {
                            $Player.GetOutOfJailFree++
                        }

                        # Advance to nearest...
                        if ($Card.Name -match 'Advance to the nearest') {
                            $Player.CurrentSpace = NextDestination -Name $($Card.Name).Trim().Split()[-1]
                            $CurrentSpace = $Spaces[$Player.CurrentSpace]
                            UpdatePlayerPieces
                            Write-Host "$($Player.Name) moved to " -NoNewline -ForegroundColor $Player.Color
                            if ($CurrentSpace.SetColor -eq 'Orange') {
                                Write-Host $SpaceColor"$($CurrentSpace.Name)"`n$ResetColor
                            } else {
                                Write-Host "$($CurrentSpace.Name)"`n -ForegroundColor $SpaceColor
                            }
                            $PayDouble = $true
                            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                            SpaceActions
                        }
                    }

                    Jail {
                        Write-Host "Go straight to jail! $($Player.Name) must roll a double to get out."`n -ForegroundColor $Player.Color
                        $Player.CurrentSpace = 10
                        UpdatePlayerPieces
                        $Player.InJail = $true
                        $Double = $false
                        if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    }

                    IRS {
                        Write-Host "$($Player.Name) must pay `$$($CurrentSpace.Rent)!"`n -ForegroundColor $Player.Color
                        $Player.Money -= $CurrentSpace.Rent
                        if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                    }

                    Null {
                        $null
                    }
                }
            }
            SpaceActions
            $BankruptcyCheck = BankruptcyCheck
            $Players = $BankruptcyCheck.RemainingPlayers
            if ($BankruptcyCheck.IsBankrupt) {
                Write-Host "$($Player.Name) has gone bankrupt!"`n -ForegroundColor $Player.Color
                $PlayerIndex = ($PlayerIndex += 1) % ($Players | Measure-Object).Count
                continue
            }
        }
    } until (!$Double -or $Player.Money -le 0 -or $Player.InJail)

    # Unmortgaging properties
    $MortgagedProperties = $Spaces | Where {$_.Owner -eq $Player.Name -and $_.Mortgaged}
    if ($MortgagedProperties) {
        foreach ($Property in $MortgagedProperties) {
            if ($Player.Money -gt ($Property.UnmortgageCost * 5)) {
                $SpaceColor = GetPropertyColor $Property.SetColor
                if ($Player.IsHuman) {
                    Write-Host "Would you like to unmortgage " -NoNewline
                    if ($Property.SetColor -eq 'Orange' -and !$psISE) {
                        Write-Host $SpaceColor"$($Property.Name)"$ResetColor -NoNewline
                    } else {
                        Write-Host "$($Property.Name)" -ForegroundColor $SpaceColor -NoNewline
                    }
                    Write-host " for `$$($Property.UnmortgageCost)? [Y/N]"
                    $Unmortgage = GetKeyPress
                } else {
                    $Unmortgage = $true
                }
                if ($Unmortgage -eq $true) {
                    Write-Host "$($Player.Name) unmortgaged " -NoNewline -ForegroundColor $Player.Color
                    if ($Property.SetColor -eq 'Orange' -and !$psISE) {
                        Write-Host $SpaceColor"$($Property.Name)"$ResetColor -NoNewline
                    } else {
                        Write-Host "$($Property.Name)" -ForegroundColor $SpaceColor -NoNewline
                    }
                    Write-Host " for `$$($Property.UnmortgageCost)"`n -ForegroundColor $Player.Color
                    $Property.Mortgaged = $false
                    $Player.Money -= $Property.UnmortgageCost
                    UpdateProperties
                }
                if (!$Testing) {Start-Sleep -Milliseconds $Pause}
            }
        }
    }

    # Trading
    $TradedProperties = @()
    $PropertiesOwned = $Spaces | Where {$_.Owner -eq $Player.Name}
    $PropertySetsOwned = $PropertiesOwned | Group {$_.SetColor}
    foreach ($ColorSet in $PropertySetsOwned) {
        $TradeTrigger = if ($ColorSet.Group[0].Type -eq 'Railroad') {2} else {1}
        if ($ColorSet.Count -eq ($ColorSet.Group[0].NumberInSet - $TradeTrigger)) {
            $PropertyTargeted = $Spaces | Where {$_.SetColor -eq $ColorSet.Group[0].SetColor -and $_.Owner -ne $Player.Name -and $_.Owner -ne 'Bank'}
            foreach ($Property in $ColorSet.Group) {
                if ($PropertyTargeted) {
                    $TradingPartner = $Players | Where {$_.Name -eq $PropertyTargeted.Owner}
                    if ($Player.Name -ne $TradingPartner.Name) {
                        $TradingPartnerPropertiesOwned = $Spaces | Where {$_.Owner -eq $TradingPartner.Name}
                        $TradingPartnerPropertySetsOwned = $TradingPartnerPropertiesOwned | Group {$_.SetColor}
                        foreach ($TradingPartnerColorSet in $TradingPartnerPropertySetsOwned) {
                            if ($TradingPartnerColorSet.Count -eq ($TradingPartnerColorSet.Group[0].NumberInSet - 1)) {
                                $TradingPartnerPropertyTargeted = $Spaces | Where {$_.SetColor -eq $TradingPartnerColorSet.Group[0].SetColor -and $_.Owner -eq $Player.Name}
                                if ($TradingPartnerPropertyTargeted -and $TradingPartnerPropertyTargeted.SetColor -ne $PropertyTargeted.SetColor) { # Make sure properties being traded aren't in same set
                                    if ($PropertyTargeted.Name -notin $TradedProperties -and $TradingPartnerPropertyTargeted.Name -notin $TradedProperties) {
                                        $TradeValueDiff = $PropertyTargeted.Cost - $TradingPartnerPropertyTargeted.Cost
                                        switch ($TradeValueDiff) {
                                            {$_ -eq 0} {
                                                Write-Host "Trade Proposal: " -NoNewline
                                                Write-Host "$($Player.Name) " -NoNewline -ForegroundColor $Player.Color
                                                Write-Host "- " -NoNewline
                                                $SpaceColor = GetPropertyColor $TradingPartnerPropertyTargeted.SetColor
                                                if ($TradingPartnerPropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($TradingPartnerPropertyTargeted.Name)"$ResetColor -NoNewline
                                                } else {
                                                    Write-Host "$($TradingPartnerPropertyTargeted.Name) " -NoNewline -ForegroundColor $SpaceColor
                                                }
                                                Write-Host "to " -NoNewline
                                                Write-Host "$($TradingPartner.Name) " -NoNewline -ForegroundColor $TradingPartner.Color
                                                Write-Host "for " -NoNewline
                                                $SpaceColor = GetPropertyColor $PropertyTargeted.SetColor
                                                if ($PropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($PropertyTargeted.Name)"`n$ResetColor
                                                } else {
                                                    Write-Host "$($PropertyTargeted.Name)"`n -ForegroundColor $SpaceColor
                                                }
                                                if ($Player.IsHuman -or $TradingPartner.IsHuman) {
                                                    Write-Host "Do you approve of this trade? [Y/N]"
                                                    $AllowTrade = GetKeyPress
                                                } else {
                                                    $AllowTrade = $true
                                                }
                                                if ($AllowTrade -eq $true) {
                                                    Write-Host "Trade completed"`n
                                                    $PropertyTargeted.Owner = $Player.Name
                                                    $TradingPartnerPropertyTargeted.Owner = $TradingPartner.Name
                                                    UpdatePropertyOwner
                                                    $TradedProperties += $PropertyTargeted.Name, $TradingPartnerPropertyTargeted.Name
                                                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                                                }
                                            }
                                            {$_ -gt 0 -and $_ -le $Player.Money} {
                                                Write-Host "Trade Proposal: " -NoNewline
                                                Write-Host "$($Player.Name) " -NoNewline -ForegroundColor $Player.Color
                                                Write-Host "- " -NoNewline
                                                $SpaceColor = GetPropertyColor $TradingPartnerPropertyTargeted.SetColor
                                                if ($TradingPartnerPropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($TradingPartnerPropertyTargeted.Name) "$ResetColor -NoNewline
                                                } else {
                                                    Write-Host "$($TradingPartnerPropertyTargeted.Name) " -NoNewline -ForegroundColor $SpaceColor
                                                }
                                                Write-Host "and `$$($TradeValueDiff) " -NoNewline
                                                Write-Host "to " -NoNewline
                                                Write-Host "$($TradingPartner.Name) " -NoNewline -ForegroundColor $TradingPartner.Color
                                                Write-Host "for " -NoNewline
                                                $SpaceColor = GetPropertyColor $PropertyTargeted.SetColor
                                                if ($PropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($PropertyTargeted.Name)"`n$ResetColor
                                                } else {
                                                    Write-Host "$($PropertyTargeted.Name)"`n -ForegroundColor $SpaceColor
                                                }
                                                if ($Player.IsHuman -or $TradingPartner.IsHuman) {
                                                    Write-Host "Do you approve of this trade? [Y/N]"
                                                    $AllowTrade = GetKeyPress
                                                } else {
                                                    $AllowTrade = $true
                                                }
                                                if ($AllowTrade -eq $true) {
                                                    Write-Host "Trade completed"`n
                                                    $PropertyTargeted.Owner = $Player.Name
                                                    $TradingPartnerPropertyTargeted.Owner = $TradingPartner.Name
                                                    UpdatePropertyOwner
                                                    $Player.Money -= $TradeValueDiff
                                                    $TradingPartner.Money += $TradeValueDiff
                                                    $TradedProperties += $PropertyTargeted.Name, $TradingPartnerPropertyTargeted.Name
                                                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                                                }
                                            }
                                            {$_ -lt 0 -and $_ -le $TradingPartner.Money} {
                                                Write-Host "Trade Proposal: " -NoNewline
                                                Write-Host "$($Player.Name) " -NoNewline -ForegroundColor $Player.Color
                                                Write-Host "- " -NoNewline
                                                $SpaceColor = GetPropertyColor $TradingPartnerPropertyTargeted.SetColor
                                                if ($TradingPartnerPropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($TradingPartnerPropertyTargeted.Name) "$ResetColor -NoNewline
                                                } else {
                                                    Write-Host "$($TradingPartnerPropertyTargeted.Name) " -NoNewline -ForegroundColor $SpaceColor
                                                }
                                                Write-Host "to " -NoNewline
                                                Write-Host "$($TradingPartner.Name) " -NoNewline -ForegroundColor $TradingPartner.Color
                                                Write-Host "for " -NoNewline
                                                $SpaceColor = GetPropertyColor $PropertyTargeted.SetColor
                                                if ($PropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($PropertyTargeted.Name) "$ResetColor -NoNewline
                                                } else {
                                                    Write-Host "$($PropertyTargeted.Name) " -NoNewline -ForegroundColor $SpaceColor
                                                }
                                                Write-Host "and `$$([Math]::Abs($TradeValueDiff))"`n
                                                if ($Player.IsHuman -or $TradingPartner.IsHuman) {
                                                    Write-Host "Do you approve of this trade? [Y/N]"
                                                    $AllowTrade = GetKeyPress
                                                } else {
                                                    $AllowTrade = $true
                                                }
                                                if ($AllowTrade -eq $true) {
                                                    Write-Host "Trade completed"`n
                                                    $PropertyTargeted.Owner = $Player.Name
                                                    $TradingPartnerPropertyTargeted.Owner = $TradingPartner.Name
                                                    UpdatePropertyOwner
                                                    $Player.Money -= $TradeValueDiff
                                                    $TradingPartner.Money += $TradeValueDiff
                                                    $TradedProperties += $PropertyTargeted.Name, $TradingPartnerPropertyTargeted.Name
                                                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                                                }
                                            }
                                        }
                                    }
                                }
                            } elseif ($TradingPartnerColorSet.Count -eq 1) {
                                $TradingPartnerPropertyTargeted = $Spaces | Where {$_.SetColor -eq $TradingPartnerColorSet.Group[0].SetColor -and $_.Owner -eq $Player.Name} | Select -First 1
                                if ($TradingPartnerPropertyTargeted -and $TradingPartnerPropertyTargeted.SetColor -ne $PropertyTargeted.SetColor) { # Make sure properties being traded aren't in same set
                                    if ($PropertyTargeted.Name -notin $TradedProperties -and $TradingPartnerPropertyTargeted.Name -notin $TradedProperties) {
                                        $TradeValueDiff = ($PropertyTargeted.Cost * 2) - $TradingPartnerPropertyTargeted.Cost
                                        switch ($TradeValueDiff) {
                                            {$_ -eq 0} {
                                                Write-Host "Trade Proposal: " -NoNewline
                                                Write-Host "$($Player.Name) " -NoNewline -ForegroundColor $Player.Color
                                                Write-Host "- " -NoNewline
                                                $SpaceColor = GetPropertyColor $TradingPartnerPropertyTargeted.SetColor
                                                if ($TradingPartnerPropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($TradingPartnerPropertyTargeted.Name) "$ResetColor -NoNewline
                                                } else {
                                                    Write-Host "$($TradingPartnerPropertyTargeted.Name) " -NoNewline -ForegroundColor $SpaceColor
                                                }
                                                Write-Host "to " -NoNewline
                                                Write-Host "$($TradingPartner.Name) " -NoNewline -ForegroundColor $TradingPartner.Color
                                                Write-Host "for " -NoNewline
                                                $SpaceColor = GetPropertyColor $PropertyTargeted.SetColor
                                                if ($PropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($PropertyTargeted.Name)"`n$ResetColor
                                                } else {
                                                    Write-Host "$($PropertyTargeted.Name)"`n -ForegroundColor $SpaceColor
                                                }
                                                if ($Player.IsHuman -or $TradingPartner.IsHuman) {
                                                    Write-Host "Do you approve of this trade? [Y/N]"
                                                    $AllowTrade = GetKeyPress
                                                } else {
                                                    $AllowTrade = $true
                                                }
                                                if ($AllowTrade -eq $true) {
                                                    Write-Host "Trade completed"`n
                                                    $PropertyTargeted.Owner = $Player.Name
                                                    $TradingPartnerPropertyTargeted.Owner = $TradingPartner.Name
                                                    UpdatePropertyOwner
                                                    $TradedProperties += $PropertyTargeted.Name, $TradingPartnerPropertyTargeted.Name
                                                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                                                }
                                            }
                                            {$_ -gt 0 -and $_ -le $Player.Money} {
                                                Write-Host "Trade Proposal: " -NoNewline
                                                Write-Host "$($Player.Name) " -NoNewline -ForegroundColor $Player.Color
                                                Write-Host "- " -NoNewline
                                                $SpaceColor = GetPropertyColor $TradingPartnerPropertyTargeted.SetColor
                                                if ($TradingPartnerPropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($TradingPartnerPropertyTargeted.Name) "$ResetColor -NoNewline
                                                } else {
                                                    Write-Host "$($TradingPartnerPropertyTargeted.Name) " -NoNewline -ForegroundColor $SpaceColor
                                                }
                                                Write-Host "and `$$($TradeValueDiff) " -NoNewline
                                                Write-Host "to " -NoNewline
                                                Write-Host "$($TradingPartner.Name) " -NoNewline -ForegroundColor $TradingPartner.Color
                                                Write-Host "for " -NoNewline
                                                $SpaceColor = GetPropertyColor $PropertyTargeted.SetColor
                                                if ($PropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($PropertyTargeted.Name)"`n$ResetColor
                                                } else {
                                                    Write-Host "$($PropertyTargeted.Name)"`n -ForegroundColor $SpaceColor
                                                }
                                                if ($Player.IsHuman -or $TradingPartner.IsHuman) {
                                                    Write-Host "Do you approve of this trade? [Y/N]"
                                                    $AllowTrade = GetKeyPress
                                                } else {
                                                    $AllowTrade = $true
                                                }
                                                if ($AllowTrade -eq $true) {
                                                    Write-Host "Trade completed"`n
                                                    $PropertyTargeted.Owner = $Player.Name
                                                    $TradingPartnerPropertyTargeted.Owner = $TradingPartner.Name
                                                    UpdatePropertyOwner
                                                    $Player.Money -= $TradeValueDiff
                                                    $TradingPartner.Money += $TradeValueDiff
                                                    $TradedProperties += $PropertyTargeted.Name, $TradingPartnerPropertyTargeted.Name
                                                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                                                }
                                            }
                                            {$_ -lt 0 -and $_ -le $TradingPartner.Money} {
                                                Write-Host "Trade Proposal: " -NoNewline
                                                Write-Host "$($Player.Name) " -NoNewline -ForegroundColor $Player.Color
                                                Write-Host "- " -NoNewline
                                                $SpaceColor = GetPropertyColor $TradingPartnerPropertyTargeted.SetColor
                                                if ($TradingPartnerPropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($TradingPartnerPropertyTargeted.Name) "$ResetColor -NoNewline
                                                } else {
                                                    Write-Host "$($TradingPartnerPropertyTargeted.Name) " -NoNewline -ForegroundColor $SpaceColor
                                                }
                                                Write-Host "to " -NoNewline
                                                Write-Host "$($TradingPartner.Name) " -NoNewline -ForegroundColor $TradingPartner.Color
                                                Write-Host "for " -NoNewline
                                                $SpaceColor = GetPropertyColor $PropertyTargeted.SetColor
                                                if ($PropertyTargeted.SetColor -eq 'Orange' -and !$psISE) {
                                                    Write-Host $SpaceColor"$($PropertyTargeted.Name) "$ResetColor -NoNewline
                                                } else {
                                                    Write-Host "$($PropertyTargeted.Name) " -NoNewline -ForegroundColor $SpaceColor
                                                }
                                                Write-Host "and `$$([Math]::Abs($TradeValueDiff)) "`n
                                                if ($Player.IsHuman -or $TradingPartner.IsHuman) {
                                                    Write-Host "Do you approve of this trade? [Y/N]"
                                                    $AllowTrade = GetKeyPress
                                                } else {
                                                    $AllowTrade = $true
                                                }
                                                if ($AllowTrade -eq $true) {
                                                    Write-Host "Trade completed"`n
                                                    $PropertyTargeted.Owner = $Player.Name
                                                    $TradingPartnerPropertyTargeted.Owner = $TradingPartner.Name
                                                    UpdatePropertyOwner
                                                    $Player.Money -= $TradeValueDiff
                                                    $TradingPartner.Money += $TradeValueDiff
                                                    $TradedProperties += $PropertyTargeted.Name, $TradingPartnerPropertyTargeted.Name
                                                    if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    # Buying houses/hotels
    $BuildableProperties = $Spaces | Where {$_.Owner -eq $Player.Name -and $_.Type -eq 'Property'}
    $BuildablePropertySets = $BuildableProperties | Group {$_.SetColor}
    foreach ($ColorSet in $BuildablePropertySets) {
        if ($ColorSet.Count -eq ($ColorSet.Group[0].NumberInSet) -and !($ColorSet.Group | Where {$_.Mortgaged})) {
            while ($Player.Money -gt ($ColorSet.Group[0].BuildingCost * 5)) {
                $BuildingPurchased = $false
                $MinHouses = ($ColorSet.Group.HouseCount | Measure-Object -Minimum).Minimum
                $MinHotels = ($ColorSet.Group.HotelCount | Measure-Object -Minimum).Minimum
                foreach ($Property in ($ColorSet.Group | Where {$_.HouseCount -eq $MinHouses -and $_.HotelCount -eq $MinHotels})) {
                    if ($Property.HouseCount -lt 4 -and $Property.HotelCount -ne 1 -and $TotalHouses -le 32) {
                        $SpaceColor = GetPropertyColor $Property.SetColor
                        if ($Player.IsHuman) {
                            Write-Host "Would you like to purchase a house for " -NoNewline
                            if ($Property.SetColor -eq 'Orange' -and !$psISE) {
                               Write-Host $SpaceColor"$($Property.Name)"$ResetColor -NoNewline
                            } else {
                               Write-Host "$($Property.Name)" -ForegroundColor $SpaceColor -NoNewline
                            }
                            Write-Host "? [Y/N]"
                            $PurchaseHouse = GetKeyPress
                        } else {
                            $PurchaseHouse = $true
                        }
                        if ($PurchaseHouse -eq $true) {
                            $Property.HouseCount++
                            $Player.HousesOwned++
                            $TotalHouses++
                            Write-Host "$($Player.Name) bought a house for " -NoNewline -ForegroundColor $Player.Color
                            if ($Property.SetColor -eq 'Orange' -and !$psISE) {
                                Write-Host $SpaceColor"$($Property.Name)"$ResetColor -NoNewline
                            } else {
                                Write-Host "$($Property.Name)" -ForegroundColor $SpaceColor -NoNewline
                            }
                            Write-Host " `($($Property.HouseCount) of 4`)"`n -ForegroundColor $Player.Color
                            $Player.Money -= $Property.BuildingCost
                            $BuildingPurchased = $true
                            UpdateProperties
                            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                        }
                    } elseif ($Property.HouseCount -eq 4 -and $Property.HotelCount -lt 1 -and $TotalHotels -le 12) {
                        $SpaceColor = GetPropertyColor $Property.SetColor
                        if ($Player.IsHuman) {
                            Write-Host "Would you like to purchase a hotel for " -NoNewline
                            if ($Property.SetColor -eq 'Orange' -and !$psISE) {
                               Write-Host $SpaceColor"$($Property.Name)"$ResetColor -NoNewline
                            } else {
                               Write-Host "$($Property.Name)" -ForegroundColor $SpaceColor -NoNewline
                            }
                            Write-Host "? [Y/N]"
                            $PurchaseHotel = GetKeyPress
                        } else {
                            $PurchaseHotel = $true
                        }
                        if ($PurchaseHotel -eq $true) {
                            $Property.HotelCount = 1
                            $Property.HouseCount = 0
                            $Player.HotelsOwned += 1
                            $Player.HousesOwned -= 4
                            $TotalHotels++
                            $TotalHouses -= 4
                            Write-Host "$($Player.Name) bought a hotel for " -NoNewline -ForegroundColor $Player.Color
                            if ($Property.SetColor -eq 'Orange' -and !$psISE) {
                                Write-Host $SpaceColor"$($Property.Name)"$ResetColor -NoNewline
                            } else {
                                Write-Host "$($Property.Name)" -ForegroundColor $SpaceColor -NoNewline 
                            }
                            Write-Host " `($($Property.HotelCount) of 1`)"`n -ForegroundColor $Player.Color
                            $Player.Money -= $Property.BuildingCost
                            $BuildingPurchased = $true
                            UpdateProperties
                            if (!$Testing) {Start-Sleep -Milliseconds $Pause}
                        }
                    }
                }
                if (!$BuildingPurchased -or $PurchaseHouse -eq $false -or $PurchaseHotel -eq $false) {
                    break
                }
            }
        }
    }

    $BankruptcyCheck = BankruptcyCheck
    $Players = $BankruptcyCheck.RemainingPlayers
    if ($BankruptcyCheck.IsBankrupt) {
        Write-Host "$($Player.Name) has gone bankrupt!"`n -ForegroundColor $Player.Color
    }
    $TradedProperties = @()
    $PlayerIndex = ($PlayerIndex += 1) % ($Players | Measure-Object).Count
    Write-Host
    # Give choice to end game if all human players have lost
    if ($BankruptcyCheck.HumanPlayersLeft -eq 0 -and ($Players | Measure-Object).Count -gt 1 -and !$AIOnlyGame) {
        Write-Host "All human players have went bankrupt. End the game? [Y/N]"
        $ExitGame = GetKeyPress
        if ($ExitGame -eq $true) {
            Write-Host "Game Over!"
            ExitGame
            Exit
        }
    }
}

Write-Host $Players[0].Name has won the game `(`$$($Players[0].Money)`)! Press SPACEBAR to exit.
if (!$psISE) {
    do {
        while ([System.Console]::KeyAvailable) {
            [System.Console]::ReadKey($true) | Out-Null
        }
        $KeyInfo = [System.Console]::ReadKey($true)
        $KeyChar = $KeyInfo.KeyChar
    } while ($KeyChar -ne ' ')
} else {
    Read-Host
}

# Close board window
ExitGame

### Dice rolling is broken