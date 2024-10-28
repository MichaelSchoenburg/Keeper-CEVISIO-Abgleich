#-----------------------------------------------------------------------------------------------------------------
#region RULES

# Rules/Conventions:

#     # This is a comment related to the following command

#     <# 
#         This is a comment related to the next couple commands (most times a caption for a section)
#     #>

# "This is a Text containing $( $variable(s) )"

# 'This is a text without variable(s)'

# Variables available in the entire script start with a capital letter
# $ThisIsAGlobalVariable

# Variables available only locally e. g. in a function start with a lower case letter
# $thisIsALocalVariable

# Assign output to $null to suppress it. E. g. $null = [System.Reflection.Assembly]::LoadWithPartialName( "System.Windows.Forms" )

#endregion RULES
#-----------------------------------------------------------------------------------------------------------------
#region INITIALIZATION
#endregion INITIALIZATION
#-----------------------------------------------------------------------------------------------------------------
#region DECLARATIONS

<# 
    Keeper
#>

$PathKeeper = "$env:USERPROFILE\Downloads\1729973991736-keeper.csv"
$CsvKeeper = Import-Csv -Path $PathKeeper -Delimiter ','
# Private Datensätze entfernen
$CsvKeeper = $CsvKeeper.Where({$_.Column7 -Like '_*'})

<# 
    CEVISIO
#>

$PathCevisio = "$env:USERPROFILE\Downloads\Ansichtsdruck - Passwörter.csv"
$CsvCevisio = Import-Csv -Path $PathCevisio -Delimiter ';'
# Nur relevante Spalten auswählen
$CsvCevisio = $CsvCevisio | Select-Object -Property "Nummer", "Bezeichnung", "Adresse", "Adresskontakt", "IT-Inventar", "Passworttyp", "Internetseite", "Benutzername", "Geräte-ID", "Passwort"

<# 
    Filter und URL ohne Pfad parsen
#>

$CsvCevisio = $CsvCevisio | Select-Object *, @{Name='Kundennummer'; Expression={[int]$_.Adresse.Split('-')[0].Trim()}}, @{Name='UrlOhnePfad'; Expression={(Select-String -InputObject $_.Internetseite -Pattern 'https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,4}').Matches.Value}}
$CsvKeeper = $CsvKeeper | Select-Object *, @{Name='UrlOhnePfad'; Expression={(Select-String -InputObject $_.Column5 -Pattern 'https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,4}').Matches.Value}}

<# 
    Allgemein
#>

$columns = $CsvKeeper[0].psobject.properties.name
$AlleDatensätzeMitId = @()
$AlleDatensätzeOhneId = @()

#endregion DECLARATIONS
#-----------------------------------------------------------------------------------------------------------------
#region FUNCTIONS
#endregion FUNCTIONS
#-----------------------------------------------------------------------------------------------------------------
#region Verarbeitung der Keeper-Datensätze

$i = 0

foreach ($e in $CsvKeeper) {
    # Kundennummer, falls vorhanden parsen
    try {
        $KundennummerString = $e.Column1.Split('|')[0].Trim()
        if ($KundennummerString -eq "") {
            $Kundennummer = $null
        } else {
            $Kundennummer = [int]$KundennummerString
        }
    }
    catch {
        # Wenn die Umwandlung zum Integer fehlgeschlagen ist, weil es sich um Text handelt
        $Kundennummer = $null
    }

    $column = $columns.Where({($e.$_ -contains "CEVISIO-Passwort-ID") -or ($e.$_ -contains "CEVISIO-Passwort-Nummer")})

    if ($column) {
        # Entferne das Column, nimm nur die Zahl und erhöhe sie um eins
        $Num = [int][string]($column -replace "[^0-9]", '') + 1
        # Hol dir den Wert aus der entsprechenden Column/Spalte
        $CevisioPasswortIdUntrimmed = $e."Column$Num"
        # Anfang und Ende trimmen
        $CevisioPasswortId = $CevisioPasswortIdUntrimmed.Trim()
    } else {
        $CevisioPasswortId = $null
    }

    # Objekt anhand der standard Felder zusammenbauen
    $psco = [PSCustomObject]@{
        ArrayIndex          = $CsvKeeper.IndexOf($e)
        # ArrayIndex          = $i
        CevisioPasswortId   = $CevisioPasswortId
        Kundennummer        = $Kundennummer
        Titel               = $e.Column2
        Benutzername        = $e.Column3
        Passwort            = $e.Column4
        URL                 = $e.Column5
        UrlOhnePfad         = $e.UrlOhnePfad
        Notizen             = $e.Column6
        GeteilterOrdner     = $e.Column7
    }

    # Falls vorhanden benutzerdefinierte Felder hinzufügen
    $ColumnsFilled = $columns.Where({$null -ne $e.$_})
    if ($ColumnsFilled.Count -gt 7) {
        # Count - 1, weil die UrlOhnePfad übersprungen werden muss
        for ($ii = 7; $ii -lt $ColumnsFilled.Count-1; $ii++) {
            $psco | Add-Member -MemberType NoteProperty -Name $ColumnsFilled[$ii] -Value $e.$($ColumnsFilled[$ii])
        }
    }

    # Objekt zum Array hinzufügen
    $AlleDatensätzeMitId += $psco

    # Zähler erhöhen
    $i++
}

<# 
    Gültige und ungültige IDs/Nummern unterscheiden
#>

$AlleDatensätzeMitGültigerId = $AlleDatensätzeMitId.Where({$_.CevisioPasswortId -match '10[0-9]{4}'})
$AlleDatensätzeMitUngültigerId = $AlleDatensätzeMitId.Where({$_.CevisioPasswortId -notmatch '10[0-9]{4}'})

<# 
    Ansichten für Troubleshooting:
    $AlleDatensätzeMitId = $AlleDatensätzeMittId.Where({$_.tId -match '10[0-2][0-9]{3}'})

    $KeeperProperties = "ArrayIndex","CevisioPasswortId","Kundennummer","Titel","Benutzername","Passwort","URL","UrlOhnePfad","Notizen","GeteilterOrdner","Column8","Column9","Column10","Column11","Column12","Column13","Column14","Column15","Column16","Column17","Column18","Column19","Column20","Column21","Column22","Column23","Column24","Column25","Column26","Column27","Column28","Column29"
    
    $AlleDatensätzeMitGültigerId | Select-Object -Property $KeeperProperties | Out-GridView
    $AlleDatensätzeMitUngültigerId | Select-Object -Property $KeeperProperties | Out-GridView
#>

#endregion Verarbeitung der Keeper-Datensätze
#-----------------------------------------------------------------------------------------------------------------
#region Match alle Datensätze, mit einer gültigen CEVISIO-Passwort-ID/-Nummer

$MatchMitGültigerIdCount = 0
foreach ($DatensatzMitGültigerId in $AlleDatensätzeMitGültigerId) {
    $Match = $CsvCevisio.Where({$_.Nummer -eq $DatensatzMitGültigerId.CevisioPasswortId})
    if ($Match.Count -eq 0) {
        # Write-Host -ForegroundColor Red "Kein Match gefunden für $($DatensatzMitGültigerId.CevisioPasswortId)"
    } else {
        # Write-Host -ForegroundColor Green "Match gefunden für ID $($Match.Nummer) mit Bezeichnung: $($Match.Bezeichnung)"
        $MatchMitGültigerIdCount++
    }
}

#endregion Match alle Datensätze, mit einer gültigen CEVISIO-Passwort-ID/-Nummer
#-----------------------------------------------------------------------------------------------------------------
#region Matche alle Datensätze, ohne eine gültige CEVISIO-Passwort-ID/-Nummer, welche einen Benutzernamen haben

$i = 0
$MatchCount = 0

foreach ($DatensatzMitUngültigerId in $AlleDatensätzeMitUngültigerId.Where({$null -ne $_.Benutzername})) {
    $Matches = $CsvCevisio.Where({
        ($_.Kundennummer -eq $DatensatzMitUngültigerId.Kundennummer) -and ($_.Benutzername -eq $DatensatzMitUngültigerId.Benutzername)
    })

    if ($DatensatzMitUngültigerId.URL -ne "") {
        $MatchesUrl = $Matches.Where({$_.URL -eq $DatensatzMitUngültigerId.URL})
        $Matches = if ($MatchesUrl.Count -eq 0) {
            $Matches.Where({$_.UrlOhnePfad -eq $DatensatzMitUngültigerId.UrlOhnePfad})
        } else {
            $MatchesUrl
        }
    }

    if ($Matches.Count -eq 0) {
        # Write-Host -ForegroundColor Red "Kein Match für `$AlleDatensätzeMitUngültigerId[$i]"
    } elseif ($Matches.Count -eq 1) {
        $MatchCount++

        $MatchTable = @(
            [PSCustomObject]@{
                Quelle        = "CEVISIO"
                Nummer        = $Matches.Nummer
                Titel         = $Matches.Bezeichnung
                Kundennummer  = $Matches.Kundennummer
                Url           = $Matches.Internetseite
                UrlOhnePfad   = $Matches.UrlOhnePfad
                Benutzername  = $Matches.Benutzername
                Passwort      = $Matches.Passwort
            },
            [PSCustomObject]@{
                Quelle        = "Keeper"
                Nummer        = $DatensatzMitUngültigerId.CevisioPasswortId
                Titel         = $DatensatzMitUngültigerId.Titel
                Kundennummer  = $DatensatzMitUngültigerId.Kundennummer
                Url           = $DatensatzMitUngültigerId.URL
                UrlOhnePfad   = $DatensatzMitUngültigerId.UrlOhnePfad
                Benutzername  = $DatensatzMitUngültigerId.Benutzername
                Passwort      = $DatensatzMitUngültigerId.Passwort
            }
        )

        # Output
        $MatchTable | Format-Table -AutoSize
        Write-Host -ForegroundColor Blue "---------------------------------------------------------------------------------"
    }
    $i++
}

#endregion Matche alle Datensätze, ohne eine gültige CEVISIO-Passwort-ID/-Nummer, welche einen Benutzernamen haben

Write-Host -ForegroundColor Yellow "Anzahl der Matches ohne gültige CEVISIO-Passwort-Nummer = $($MatchCount)"
Write-Host -ForegroundColor Yellow "Anzahl der Matches mit gültiger CEVISIO-Passwort-Nummer = $($MatchMitGültigerIdCount)"
Write-Host -ForegroundColor Yellow "Anzahl der Keeper-Datensätze ohne Match = $($CsvKeeper.Count - ($MatchMitGültigerIdCount + $MatchCount))"
