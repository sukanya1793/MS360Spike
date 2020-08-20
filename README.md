# MS360Spike

This project uses MS Graph SDK to hit Graph APIs by generating an OAUTH2 access token. 
You could do one of the following :- 
1. Display your access token
2. Display your Outlook calendar events
3. Display your license information along with their name and provisioning status 

### Prerequisites

1. A Microsoft Account
2. .NET Core SDK installed on your machine

#### STEP 1: Clone the project

```git clone https://github.com/sukanya1793/MS360Spike.git```

#### STEP 2: Follow below document and create azure applicaiton
```https://docs.microsoft.com/en-us/graph/tutorials/dotnet-core?tutorial-step=2```

#### STEP 3: Follow below document and set appid & scope
```https://docs.microsoft.com/en-us/graph/tutorials/dotnet-core?tutorial-step=3```

#### STEP 4: Build the project

```dotnet build```

#### STEP 5: Run the project

```dotnet run```

Click on the link and sign in to your Microsoft account using the code provided, after which 3 options will be displayed. 
#### Note: The Calendar events will be displayed only if you have Outlook enabled for your mail ID.

#### Troubleshooting
1. Make sure redirect uri for application is set with below URI's
https://login.microsoftonline.com/common/oauth2/nativeclient 
https://login.live.com/oauth20_desktop.srf (LiveSDK) 
msal79f8403b-cec8-4bb2-b134-17a9005482fe://auth (MSAL only)
 
