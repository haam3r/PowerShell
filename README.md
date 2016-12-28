# PowerShell
Collection of PowerShell scripts

## Notes

### Disable PowerShell Executionpolicy

**NB!Has to be run from elevated prompt**

```powershell
powershell.exe -NoP -NonI -W Hidden -Enc UwBlAHQALQBFAHgAZQBjAHUAdABpAG8AbgBQAG8AbABpAGMAeQAgAC0ARQB4AGUAYwB1AHQAaQBvAG4AUABvAGwAaQBjAHkAIABVAG4AcgBlAHMAdAByAGkAYwB0AGUAZAAgAC0ARgBvAHIAYwBlAA==
```
Enc portion is base64 encoded and contains:
```powershell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
```

### Run powershell functions from Python

Works on Python 2.7.10

```powershell
 p = subprocess.Popen(["powershell.exe", ". C:\\Temp\\test.ps1;", "test -Text 'test2'"], stdout=sys.stdout)
 ```
