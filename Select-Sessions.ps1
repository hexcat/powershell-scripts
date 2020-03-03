function Select-Sessions
{
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true,
        ValueFromPipeline=$true)]
        [System.Diagnostics.Eventing.Reader.EventLogRecord]
        $Event,

        [parameter(Mandatory=$true)]
        [System.DateTime[]]
        $Dates,

        [parameter(Mandatory=$false)]
        [uint32]
        $MaximumSessions = 200
    )


    Begin
    {
        $CurrentSessions = @{}
        $Dates = $Dates | Sort-Object
        $TotalSeconds = $null
        $ReportsFinished = 0

        $Report = @{}
        foreach ($Date in $Dates) {
            $Report[$Date] = New-Object System.Collections.ArrayList
        }

        function Read-EventDetail ($Event, $Detail) {
            $Regex = $Detail +':\s*(?<detail>\S+|)(\r?\n|$)'

            if ($Event.Message -match $Regex) {
                return $Matches['detail']
            }

            return ''
        }
        
        function Read-EventDetails($Event, [String[]]$DetailStrings) {
            $DetailsOrder = 
                'User', 'LogonId', 'LogonType', 'SourceAddress'

            $Details = @{}
            for ($i = 0; $i -lt $DetailsOrder.Length; $i++) {
                if ($DetailStrings[$i]) {
                    $Details[$DetailsOrder[$i]] = Read-EventDetail $Event $DetailStrings[$i]
                } else {
                    $Details[$DetailsOrder[$i]] = '-'
                }
            }
            $Details['LogonId'] = $Details['LogonId'].ToUpperInvariant()
            $Details['EventTime'] = $Event.TimeCreated

            return $Details
        }

        function Add-Session ($Event, [String[]]$Details, [String]$Status = 'Активно', [switch]$InHindsight, [switch]$Force) {
            $Properties = (Read-EventDetails $Event $Details)
            $Properties['Status'] = $Status

            if (-not $InHindsight) {
                if (-not $CurrentSessions.ContainsKey($Properties['LogonId']) -or $Force) {
                    $CurrentSessions[$Properties['LogonId']] = (New-Object psobject -Property $Properties)
                }
            } else {
                foreach ($Date in $Dates) {
                    if ($Event.TimeCreated -ge $Date) {
                        $Report[$Date].Add( (New-Object psobject -Property $Properties) ) | Out-Null
                    }
                }
            }
        }

        function Remove-Session ($Event, [String[]]$Details, [switch]$Keep) {
            $LogonId = (Read-EventDetail $Event $Details[1]).ToUpperInvariant()

            if ($CurrentSessions.ContainsKey($LogonId)) {
                if (-not $Keep) {
                    $CurrentSessions.Remove($LogonId)
                } else {
                    $CurrentSessions[$LogonId].Status = 'Завершено'
                }
            } else {
                Add-Session -InHindsight -Status 'Неизвестно (найден выход)' $Event $Details
            }
        }

        function Switch-Session ($Event, [String[]]$Details, [switch]$Connected) {
            $LogonId = (Read-EventDetail $Event $Details[1]).ToUpperInvariant()
            $SourceAddress = Read-EventDetail $Event $Details[3]

            if ($CurrentSessions.ContainsKey($LogonId)) {
                if ($Connected) {
                    $CurrentSessions[$LogonId].Status = 'Активно'
                    $CurrentSessions[$LogonId].SourceAddress = $SourceAddress
                } else {
                    $CurrentSessions[$LogonId].Status = 'Отключено'
                }
            } else {
                if ($Connected) {
                    Add-Session $Event $Details
                    Add-Session -InHindsight -Status 'Отключено (найдено подключение)' $Event $Details
                } else {
                    Add-Session -Status 'Отключено' $Event $Details
                    Add-Session -InHindsight -Status 'Активно (найдено отключение)' $Event $Details
                }
            }
        }
    }

    Process
    {
        if (-not $TotalSeconds) {
            $TotalSeconds = ($Dates[-1] - $Event.TimeCreated).TotalSeconds
            $FirstEventTime = $Event.TimeCreated
        } else {
            if ($ReportsFinished -lt $Dates.Count) {
                Write-Progress `
                    -Activity "Формирование отчета на момент: $($Dates[$ReportsFinished])" `
                    -CurrentOperation $Event.TimeCreated `
                    -PercentComplete ( (($Event.TimeCreated - $FirstEventTime).TotalSeconds * 100) / $TotalSeconds ) `
            } else {
                Write-Progress `
                    -Activity "Обработка остальных событий" `
                    -CurrentOperation $Event.TimeCreated `
                    -PercentComplete (-1)
            }
        }

        if ($CurrentSessions.Count -gt $MaximumSessions) {
            $TimeLimit = $Event.TimeCreated.AddDays(-7)
            $SessionsToRemove = $CurrentSessions.Values | Where-Object { $_.Status -eq 'Завершено' -and $_.Time -lt $TimeLimit }
            foreach ($Session in $SessionsToRemove) {
                $CurrentSessions.Remove($Session)
            }

            if ($CurrentSessions.Count -gt $MaximumSessions) {
                throw 'Слишком много незавершенных сессий.' +
                      'Увеличьте параметр MaximumSessions (по умолчанию равен 200)'
            }
        }

        while ($ReportsFinished -lt $Dates.Count -and $Event.TimeCreated -ge $Dates[$ReportsFinished]) {
            $Report[$Dates[$ReportsFinished++]] = [System.Collections.ArrayList](
                $CurrentSessions.Values | Where-Object { $_.Status -ne 'Завершено' } | ForEach-Object { $_.PsObject.Copy() })
        }

        Switch ($Event.Id) {
            # Успешный вход в систему
            # Успешный сетевой вход в систему
            {$_ -eq 528 -or $_ -eq 540} {
                Add-Session -Force $Event 'Пользователь', 'ИД входа', 'Тип входа', 'Адрес сети источника'
            }

            # Выход пользователя из системы
            538 {
                Remove-Session $Event 'Пользователь', 'Код входа', 'Тип входа', $null
            }

            # Выход, вызванный пользователем
            551 {
                Remove-Session -Keep $Event 'Пользователь', 'Код входа', $null, $null
            }

            # Присвоение специальных прав для нового сеанса входа
            576 {
                Add-Session $Event 'Пользователь', 'Код входа', $null, $null
            }


            # Сеанс подключен к станции
            682 {
                Switch-Session -Connected $Event 'Имя пользователя', 'Код входа', $null, 'Адрес клиента'
            }

            # Сеанс отключен от станции
            683 {
                Switch-Session $Event 'Имя пользователя', 'Код входа', $null, 'Адрес клиента'
            }

            # Игнорируются
            # 552 Попытка входа с явным указанием учетных данных
        }


    }

    End
    {
        while ($ReportsFinished -lt $Dates.Count) {
            $Report[$Dates[$ReportsFinished++]] = [System.Collections.ArrayList]($CurrentSessions.Values |
                Where-Object { $_.Status -ne 'Завершено' } | ForEach-Object { $_.PsObject.Copy() })
        }

        $MergedReport = New-Object System.Collections.ArrayList
        foreach ($Date in $Dates) {
            $MergedReport.AddRange( ($Report[$Date] | Select-Object @{ N = 'Time' ; E = {$Date} },*) ) | Out-Null
        }

        return $MergedReport
    }

}
