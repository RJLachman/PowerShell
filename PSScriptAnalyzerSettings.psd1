@{
    # Alleen fouten en waarschuwingen worden gerapporteerd (en laten CI falen).
    Severity     = @('Error', 'Warning')

    # Alle standaardregels zijn actief, behalve:
    ExcludeRules = @(
        # Dit zijn interactieve beheerscripts; gekleurde console-uitvoer is bewust.
        'PSAvoidUsingWriteHost'
    )

    Rules        = @{
        PSUseCompatibleSyntax      = @{
            Enable         = $true
            TargetVersions = @('5.1', '7.0')
        }
        PSAvoidUsingCmdletAliases  = @{
            Enable = $true
        }
    }
}
