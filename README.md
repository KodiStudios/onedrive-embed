# onedrive-embed

Using this script involves first getting OneDrive Token, and
using it in script.

Essentially, this script uses permissions of OneDrive WebApp.

## Retrieve Token

Token is needed to access OneDrive. Let's retrieve it.

Navigate to Microsoft Graph Explorer:  
http://aka.ms/ge

Click on **Getting Started** > **list items in my drive**  
Click **Run query** button

If query fails, click on "Modify Permissions"
Add permission: "Files.Read"

Ensure Query Executed Correctly  
Click **Access Token** button

Copy token value to clipboard, this will be your {token-value} in next step.

## Execute

Execute:

```Cmd
node index.ts --directory {your-directory} --token {token-value}
```
