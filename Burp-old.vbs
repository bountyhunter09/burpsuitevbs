Set objShell = CreateObject("WScript.Shell")

javaCmd = "java -Xbootclasspath/p:""C:\Program Files\Java\jdk1.8.0_202\bin\burp-loader-keygen.jar"" -jar ""C:\Program Files\Java\jdk1.8.0_202\bin\burpsuite_pro_v1.7.37.jar"""

' Run the Java command
objShell.Run javaCmd, 0, False

' Wait for a few seconds
WScript.Sleep 5000  ' 5000 milliseconds = 5 seconds
