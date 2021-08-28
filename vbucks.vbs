option explicit

dim var
        var = InputBox("Type amount:", "Get free vbucks now!")

        if var = "" then
                Msgbox "Write something!"
        else
                MsgBox "You selected: " & var & " vbucks.", vbInformation   
        end if
dim rebcsd
rebcsd = InputBox("Type username:", "Get free vbucks now!")
if rebcsd = "" Then
        MsgBox "Write something!"
else
        MsgBox "Your username is: " & rebcsd, vbInformation
end if
dim vsdc
vsdc = InputBox("Confirm adding?", "Get free vbucks now!")
if vsdc = "yes" Then
        MsgBox "Vbucks added. Please check.", vbInformation
Else
        MsgBox "Please type yes or no!", vbInformation
end if
if vsdc = "no" Then
        MsgBox "Vbucks aren't added."
else
        MsgBox "Please type yes or no!"
end if


