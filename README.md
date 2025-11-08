# Office2019-Activator

Install and activate Office 2019 for FREE legally using Volume license

**[Link to download]**(https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img)

## Step 1: Open Command Prompt in Administrator Mode

Open the Command Prompt with administrator privileges.

Here’s how you can **open Command Prompt (cmd) with administrator privileges** in Windows 👇

### 🪟 **Method 1: Using the Start Menu**

1. Click **Start** (or press the **Windows key**).
2. Type **cmd** or **Command Prompt**.
3. Right-click on **Command Prompt** from the search results.
4. Select **“Run as administrator.”**
5. If prompted by **User Account Control (UAC)**, click **Yes**.

### ⚡ **Method 2: Using the Run Dialog**

1. Press **Windows + R** to open the Run box.
2. Type:

   ```
   cmd
   ```
3. Then press **Ctrl + Shift + Enter** instead of just Enter.
   → This directly opens Command Prompt as Administrator.
4. Click **Yes** on the UAC prompt.

### 💡 **Method 3: Using Task Manager**

1. Press **Ctrl + Shift + Esc** to open Task Manager.
2. Click **File → Run new task**.
3. Type `cmd`.
4. Check the box **“Create this task with administrative privileges.”**
5. Click **OK**.

---

## Step 2: Locate the Office Installation Directory

Navigate to the directory where Office is installed on your PC. Run the following commands:

```cmd
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
```

If Office is installed in the `ProgramFiles` folder, the path will be either `%ProgramFiles%\Microsoft Office\Office16` or `%ProgramFiles(x86)%\Microsoft Office\Office16`, depending on the architecture of your Windows OS. If you are unsure, run both commands. One will fail and display an error message.

## Step 3: Convert Retail License to Volume License

Convert your retail license to a volume license by running the following command:

```cmd
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
```

## Step 4: Activate Office Using KMS Client Key

Ensure your PC is connected to the internet, then run the following commands to activate Office:

```cmd
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6MWKP
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:23.226.136.46
cscript ospp.vbs /act
```
