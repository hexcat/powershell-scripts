function Get-SessionsSummary ($Report, [string]$LogonType) {
    $Summary = @{}
    $ReducedReport = $Report | Where-Object { $_.LogonType -eq $LogonType } | Sort-Object Time,User -Unique

    foreach ($Session in $ReducedReport) {
        if (-not $Summary.ContainsKey($Session.User)) {
            $Summary[$Session.User] = New-Object psobject -Property @{
                User = $Session.User
                Active = 0;
                Disconnected = 0;
                Finished = 0;
                Unknown = 0;
                Total = 0
            }
        }

        Switch ($Session.Status) {
            {$_ -eq 'Активно' -or $_ -eq 'Активно (найдено отключение)'} {
                $Summary[$Session.User].Active++
            }

            {$_ -eq 'Отключено' -or $_ -eq 'Отключено (найдено подключение)'} {
                $Summary[$Session.User].Disconnected++
            }

            'Завершено' {
                $Summary[$Session.User].Finished++
            }

            'Неизвестно (найден выход)' {
                $Summary[$Session.User].Unknown++
            }
        }

        $Summary[$Session.User].Total++
    }

    $MergedSummary = New-Object System.Collections.ArrayList
    foreach ($User in $Summary.Keys) {
        $MergedSummary.Add( ($Summary[$User]) ) | Out-Null
    }

    return $Summary.Values
}
