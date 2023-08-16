Set objShell = CreateObject("WScript.Shell")

javaCmd = "C:\Program Files\Java\jdk-17\bin\java.exe" -jar "C:\your_path\your_path\your_path\burpsuite_community_v2023.9.1.jar"

' Run the Java command
objShell.Run javaCmd, 0, False

' Wait for a few seconds
WScript.Sleep 5000  ' 5000 milliseconds = 5 seconds
