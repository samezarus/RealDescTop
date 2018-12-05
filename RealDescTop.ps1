# Скрипт создаёт ярлык на диск С (реального компьютера), коорый подключается пользователю в терминальной сессиии

Clear-Host
# ------------------------------------------------------------------------------------------------

$user = $env:UserName
$linkName = 'Стол раба'
$path1 = '\\tsclient\c\Users\'+$user+'\desktop\'
$path2 = 'c:\Users\'+$user+'\desktop\'+$linkName+'.lnk'

if (Test-Path -Path $path1)
{
    # Ярлык на рабочий стол
    $shell = New-Object -comObject Wscript.Shell
    $shortcut = $shell.CreateShortcut($path2)
    $shortcut.TargetPath = $path1
    $shortcut.Save()

    # В раздел "Быстрый доступ" (win10/srv16)
    $o = new-object -com shell.application
    $o.Namespace($path1).Self.InvokeVerb('pintohome')
}
