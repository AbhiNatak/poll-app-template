# Prerequisites
To begin, you will need:
* A Microsoft 365 subscription
* A team with the users who will be sending Polls using this app. (You can add and remove team members later!)
* A copy of the Poll app GitHub repo (https://github.com/OfficeDev/microsoft-teams-poll-app)


# Step 1: Create your Poll app

To create the Teams Poll app package:
1. Make sure you have cloned the app repository locally.
1. Open the `actionManifest.json` file in a text editor.
1. Change the placeholder fields in the manifest to values appropriate for your organization.
    * package.id - A unique identifier for this app in reverse domain notation. E.g: com.example.myapp. (Max length : 64)
    * developer.name ([What's this?](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema#developer))
    * developer.websiteUrl
    * developer.privacyUrl
    * developer.termsOfUseUrl

1. Create a ZIP package with. 
    * Name this package `poll-output.zip`, so you know that this is the Poll app.
    * Make sure that there is no change to file structure of the ZIP package, with no new nested folders.



# Step 2: Deploy app to your M365 subscription
1. Open Command Line on your machine.
1. Traverse to the app package ZIP with the name `poll-output.zip`
1. Run the following command to download all the dependent files mentioned in package.json file of the app package. 

    **```npm install```**
1. Once the dependent files are downloaded, run the following command to deploy the app package to M365 Action service.

    **```npm run create```**
1. When prompted, log in to your M365 subscription.
1. M365 Action service will programmatically create an "AAD Custom app" in your tenant and create a “M365 subscription backed Bot” to power the Poll message extension app in Teams.
1. M365 Action service will generate a Poll Teams app zip file with the name `poll-teams-upload.zip` in the same directory as your cloned app repository locally.


# Step 3: Run the app in Microsoft Teams

If your tenant has sideloading apps enabled, you can install your app by following the instructions [here](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload#load-your-package-into-teams).

You can also upload it to your tenant's app catalog, so that it can be available for everyone in your tenant to install. See [here](https://docs.microsoft.com/en-us/microsoftteams/tenant-apps-catalog-teams).

Upload the generated Poll Teams app zip file (the `poll-teams-upload.zip` package) to your channel, chat, or tenant’s app catalog.

# Step 4: Update your Poll Teams app

If you want to update the existing Poll Teams app with latest functionality -
1. Make sure you have cloned the latest app repository locally.
1. Open the `actionManifest.json` file in a text editor.
1. Change the placeholder fields in the manifest with existing values in your Poll Teams app.
1. Create a ZIP package.
1. Run the following command to update your Poll Teams app with the latest bits of code.
    
    **```npm run update```**
1. When prompted, log in to your M365 subscription
1. Poll app on Teams automatically gets updated to the latest version.


# Scripts

## ```npm run build```
Build the app and generate output folder.

## ```npm run start```
Build the app and generate output folder along with map files for all JS. Also watch the input files and rebuild if there is any change.

## ```npm run zip```
Zip the content of output folder and create file `ActionPackage.zip`.

## ```npm run create```
Upload the `ActionPackage.zip` to ActionPlatfrom and generate `<packageId>.zip` file in output folder.

## ```npm run update```
Upload the `ActionPackage.zip` to ActionPlatfrom.

## ```npm run inner-loop```
Replace `<packageId>` with actual package id mentioned in action manifest in package.json before run this command. This command is useful for devlopment as the package is serve from output folder instead of action service.