
Set oShell = CreateObject("WScript.Shell")
cmd = "cmd /c python -m streamlit run label_app.py --server.port 8501 --server.headless false"
oShell.Run cmd, 0, False
WScript.Sleep 2000
oShell.Run "http://localhost:8501/", 0, False
