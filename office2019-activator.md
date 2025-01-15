# Office2019-Activator

Install and activate Office 2019 for FREE legally using Volume license

## Step 1: Open Command Prompt in Administrator Mode

1. Open the Command Prompt with administrator privileges.

## Step 2: Locate the Office Installation Directory

2. Navigate to the directory where Office is installed on your PC. Run the following commands:

    ```cmd
    cd /d %ProgramFiles%\Microsoft Office\Office16
    cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
    ```

    If Office is installed in the `ProgramFiles` folder, the path will be either `%ProgramFiles%\Microsoft Office\Office16` or `%ProgramFiles(x86)%\Microsoft Office\Office16`, depending on the architecture of your Windows OS. If you are unsure, run both commands. One will fail and display an error message.

## Step 3: Convert Retail License to Volume License

3. Convert your retail license to a volume license by running the following command:

    ```cmd
    for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
    ```

## Step 4: Activate Office Using KMS Client Key

4. Ensure your PC is connected to the internet, then run the following commands to activate Office:

    ```cmd
    cscript ospp.vbs /setprt:1688
    cscript ospp.vbs /unpkey:6MWKP >nul
    cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
    cscript ospp.vbs /sethst:107.175.77.7
    cscript ospp.vbs /act
