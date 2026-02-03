# onedrive-embed

## Retrieve Token

Token is needed to access OneDrive. Let's retrieve it.

Navigate to Microsoft Graph Explorer:  
http://aka.ms/ge

Click on Getting Started > list items in my drive
Click `Run query` button
Ensure Query Executed Correctly
Click `Access Token` button
Copy token value to clipboard, this will be your {token-value} in next step.

## Execute

Execute:

```Cmd
node index.ts --directory {your-directory} --token {token-value}
```
