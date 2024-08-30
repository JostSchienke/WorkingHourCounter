# Working Hour Counter

this is just a Simple litte Script to help counting your Work hours.

you can either logon or logoff (with a reason), or just export your working hours.

## commands
            "-logon               =>      Logs your start time",
            "-logoff              =>      Logs you off with the argument 'hours full'",
            " -> LOGOFF-REASON    =>      Adds a reason for the Logoff to the logoff",
            "-get-times           =>      A Command for Exporting the Times",
            " -> Date DD.MM.YYYY  =>      Exports the time of a specific day",
            " -> Month (MM)       =>      Exports the times of that specific month",
            " -> Full-Export      =>      Exports all entries as an xlsx file",
            "-help                =>      Shows all commands"

## Powershell script

their is also a PS script that can be added to you $PROFILE file, to add a Command instead of using the python ./hourcounter.
the command is "imworking" but this can be changed by changing the PS script.
