Poll app is built on M365 Action platform which uses the M365 Action service to host the app, and the M365 Action JavaScript Client SDK that makes it easy for the app to interact with the M365 Action service.

## Poll App package

The app package contains the manifest JSON file, assets file includes app icons and string files, different app views for Poll creation, Poll response, Poll results. Along with these, the files folder contains M365 Actions SDK, and the dependent Node packages.
The Poll message extension app is implemented as a React application, the Poll app views are built using UI components from [NorthStar UI](https://github.com/stardust-ui/react) and [Office UI Fabric React](https://github.com/OfficeDev/office-ui-fabric-react).

## M365 Action SDK

The M365 Action JavaScript client SDK makes it easy for developers to interact with M365 Action service. Broadly APIs in the ActionSDK.js enable the following:

1. **Poll Creation APIs:** Tapping on the App's icon in palette launches the Poll creation flow. Using these APIs you can initialize a form object, manipulate it, and submit it as a request.
1. **Poll Response APIs:** Tapping on the respond button of adaptive card launches its response flow. Using these APIs you can get the associated form object, all the previous responses, and submit a new response.
1. **Poll Result APIs:** Tapping on the result button of adaptive card launches its view result flow. Using these APIs you can get the associated form object, all the aggregated responses by the participants, and choose to close the form so that further responses are not allowed.

## M365 Action Service 

M365 Action service is an extension to Kaizala Action service to all M365 customers. The Poll app is hosted on M365 Action service and comes with the following features:
* M365 Action service supports High availability and Data redundancy.
* M365 Action service has compliant storage on par with Teams including GDPR compliance, Go-local support for Data residency in the following countries. 
