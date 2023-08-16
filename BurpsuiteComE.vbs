Set objShell = CreateObject("WScript.Shell")

javaPath = "C:\Program Files\Java\jdk-17\bin\java.exe"
jarPath = "D:\burpsuite_all_shorcuts_and_vbs_scripts\burpsuite_community_v2023.9.1.jar"

command = """" & javaPath & """ --add-opens=java.desktop/javax.swing=ALL-UNNAMED --add-opens=java.base/java.lang=ALL-UNNAMED --add-opens=java.base/jdk.internal.org.objectweb.asm=ALL-UNNAMED --add-opens=java.base/jdk.internal.org.objectweb.asm.tree=ALL-UNNAMED --add-opens=java.base/jdk.internal.org.objectweb.asm.Opcodes=ALL-UNNAMED -jar """ & jarPath & """"

objShell.Run command, 0, False ' Use "False" to close the command prompt window