function Select-1CShortcuts
{
    <#
    .SYNOPSIS
    Выбирает ярлыки, указывающие на 1С, и возвращает их свойства и параметры запуска.

    .EXAMPLE
    Get-ChildItem 'C:\Users\Ivanov\Desktop\*.lnk' | Select-1CShortcuts

    ShortcutDir  : C:\Users\Ivanov\Desktop
    ShortcutName : 1C.lnk
    Dir          : C:\Program Files\1cv82\common
    Pass         : Пароль
    Name         : 1cestart.exe
    Login        : Пользователь
    ShortcutPath : C:\Users\Ivanov\Desktop\1C.lnk
    Path         : C:\Program Files\1cv82\common\1cestart.exe
    Args         : enterprise /S10.0.0.1/database /NПользователь /PПароль /AppAutoCheckMode /ClearCache
    Server       : 10.0.0.1/database


    .EXAMPLE
    Get-ChildItem 'C:\Users\a*\Desktop\*.lnk' | Select-1CShortcuts | Format-Table Login, Pass

    Login                                                                                      Pass
    -----                                                                                      ----
    Пользователь1                                                                              Пароль1
    Пользователь2                                                                              Пароль2
    ...                                                                                        ...

    #>

    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true,
        ValueFromPipeline=$true)]
        [System.IO.FileInfo]
        $ShortcutFile
    )

    Begin
    {
        $Shell = New-Object -ComObject WScript.Shell -Strict # COM-объект, используется для "открытия" ярлыка
        $1CNames = "1cestart.exe", "1cv8.exe" # список названий клиента 1С

        # Функция для поиска в ярлыке имени пользователя 1С и его пароля с помощью регулярных выражений
        # Если имя или пароль указано несколько раз, выбирается последний вариант (клиент 1С делает так же)
        function Read-Parameter ($str, $Param) {
            $Regex = '(?x)' + # free-spacing mode (возможность оставлять комментарии)
                $Param +
                '
                    \s* # пробелы между ключом параметра и самим параметром
                    (
                        # используется именованая группа "key"
                       "(?<key>[^"]+)" | # параметр в кавычках
                        (?<key>\S+)      # параметр без кавычек
                    )
                '

            $Match = ($str | Select-String -Pattern $Regex -AllMatches)
            if ($Match) { return $Match.Matches[-1].Groups['key'].Value }

            return ''
        }
    }

    Process
    {
        $Shortcut = $Shell.CreateShortcut($ShortcutFile.FullName) # "открываем" ярлык (закрывать не нужно)
        $TargetName = $Shortcut.TargetPath | Split-Path -Leaf -ErrorAction SilentlyContinue # при ошибке из-за "особого" ярлыка продолжаем работу
        if (-not $? -or $1CNames -notcontains $TargetName) {
            return # отбрасываем ярлык, если он "особый" (не смогли прочесть) или не указывает на клиент 1С
        }

        # Чтение свойств ярлыка
        $Properties = @{
            ShortcutPath = $ShortcutFile.FullName             # полный путь к ярлыку
            ShortcutDir = $ShortcutFile.DirectoryName         # папка, содержащая ярлык
            ShortcutName = $ShortcutFile.Name                 # имя ярлыка
            Path = $Shortcut.TargetPath                       # полный путь к клиенту 1С
            Dir = $Shortcut.TargetPath | Split-Path -Parent   # папка с клиентом 1С
            Name = $TargetName                                # исполняемый файл клиента 1С
            Args = $Shortcut.Arguments                        # аргументы для клиента 1С, прописанные в ярлыке
        }
        # Поиск в ярлыке имени пользователя 1С и его пароля с помощью регулярных выражений
        # Если имя или пароль указано несколько раз, выбирается последний вариант (клиент 1С делает так же)
        $Properties.Login = Read-Parameter $Shortcut.Arguments '/N'
        $Properties.Pass = Read-Parameter $Shortcut.Arguments '/P'
        $Properties.Server = Read-Parameter $Shortcut.Arguments '/S'

        # Упаковываем свойства ярлыка в объект и возвращаем его
        $PSCmdlet.WriteObject( (New-Object PSObject -Property $Properties) )
    }

    End
    {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null # COM-объект больше не нужен, избавляемся
    }
}
