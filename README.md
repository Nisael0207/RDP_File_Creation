# RDP_File_Creation
This gui allows you to create an .rdp and sends it as mail to a user. Can also be used with remotedesktopserver

Things that need to be done:
You need to look over the variables from Line 5 to 14. The localgroup doesnt have to be changed, you maybe got groups in your
active directory for the remotedesktopuser, if you dont have thoose groups you can insert $null (as you see in line 7). For teh $rdsgateway you insert
your RD-Server. In line 11 you need your exchange / SMTP-Server and in line 12 the mail you want to send the mails from. If you dont want to send any mail leave
the checkbox in the gui unchecked.
In line 13 and 14 you can change the subject of the mail ($subject2) and the body as html code.
