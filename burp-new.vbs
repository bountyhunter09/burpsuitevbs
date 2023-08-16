Set objShell = CreateObject("WScript.Shell")

javaPath = """C:\Program Files\Java\jdk-17\bin\java.exe"""
javaArguments = "--add-opens=java.desktop/javax.swing=ALL-UNNAMED --add-opens=java.base/java.lang=ALL-UNNAMED --add-opens=java.base/jdk.internal.org.objectweb.asm=ALL-UNNAMED --add-opens=java.base/jdk.internal.org.objectweb.asm.tree=ALL-UNNAMED --add-opens=java.base/jdk.internal.org.objectweb.asm.Opcodes=ALL-UNNAMED"
javaAgent = "-javaagent:D:/burpsuite_all_shorcuts_and_vbs_scripts/Burp/BurpLoaderKeygen.jar"
jarPath = """D:\burpsuite_all_shorcuts_and_vbs_scripts\Burp\burpsuite_pro_v2023.8.jar"""

command = javaPath & " " & javaArguments & " " & javaAgent & " -noverify -jar " & jarPath

objShell.Run command, 0, True ' Change the window style to hidden (0)

' Close the script
WScript.Quit
