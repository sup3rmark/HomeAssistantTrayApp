# HomeAssistantTrayApp
 An application that sits in the Windows system tray to allow the user to trigger HomeAssistant automations.
Requirements:
- HomeAssistant
- CredentialManager module
- Server address and API token stored in Windows Credential Manager

Instructions:

1. Download HomeAssistantTrayApp.ps1 to your computer from the link above.
2. Install the CredentialManager module (`Install-Module CredentialManager` should do the trick)
3. Open a Powershell prompt and run the command

        Set-ExecutionPolicy Unrestricted

4. Store the credentials of the email address you want to send the email from in Windows Credential Manager using the CredentialManager module:

```powershell
Install-Module CredentialManager
New-StoredCredential -Target HomeAssistant -UserName [server address] -Password [API token] -Type Generic -Persist LocalMachine
```
*note: the server address should be in the full URL format (i.e. "https://my.homeassisantserver.com")*

*another note: "a [long-lived access token](https://developers.home-assistant.io/docs/auth_api/#long-lived-access-token) can be created using the UI tool located at the bottom of the user's Home Assistant profile page")*

5. Run the script! You can run from a Powershell prompt, or by right-clicking and selecting *Run*.

Default behavior:

    .\HomeAssistantTrayApp.ps1

There are no parameters for this script. The server address and token should be stored in Credential Manager. All automations are retrieved when the script is initially run and stored until the application is closed from the context menu. They are displayed in alphabetical order, and you can refresh the list by relaunching the script (right-click the icon and select "Relaunch Tray App").

Please let me know if you have any questions, comments, or suggestions!
