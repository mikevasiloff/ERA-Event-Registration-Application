# E.R.A (Event Registration Application)
A tool for organizers to create and manage events for airforce classes and training in SharePoint Online.

# Build
1. Clone the source code
2. Install the libraries
```js
npm i
```
3. Link TypeScript
```js
npm link typescript
```
4. Build the Solution
```js
npm run all
```
# Installation and Deployment
1) Go to Site content -> Site Assets where the app shall live
2) Create a new folder and it name it "Event-Registration"
![image](https://user-images.githubusercontent.com/84786032/156636796-21e15b3d-6760-46c3-bdb2-e9f75cacd5f8.png)
3) Click inside of the folder and upload the following: <br/>
 ![image](https://user-images.githubusercontent.com/84786032/156637037-fc5df5ca-4287-477b-acc1-f645db54481e.png)
4) Click inside of the Event Registration folder and upload the following files accompanying this document. 
![image](https://user-images.githubusercontent.com/84786032/156637482-0f115438-86ae-4c20-aa2f-9471fee14c6a.png)
5) Navigate to Site Pages <br/>
![image](https://user-images.githubusercontent.com/84786032/156637651-1ab7203d-8a07-44d8-b3e4-bde401a1d6e0.png)
7) Click New -> Web Part Page
![image](https://user-images.githubusercontent.com/84786032/156637716-1f2d6d00-df90-4a22-b683-ddc6b23b7e82.png)
9) Type the name of the page before ".aspx" (era is recommended) and select "Full Page, Vertical" for layout template. <br/>
![image](https://user-images.githubusercontent.com/84786032/156637835-60b3fdaa-68e4-46bf-a3f5-cbab74d3bbc8.png)
10) Exit out of the current tab and return to “Site Contents” -> “Site Assets”.
![image](https://user-images.githubusercontent.com/84786032/156638556-293a85b5-bc4d-4b6e-b618-752e66fe7930.png)
11)	Move the “.aspx” file created in Step 9) to the Event-Registration folder created in 
Step 2).
![image](https://user-images.githubusercontent.com/84786032/156638909-7bbc8522-97e0-432f-b7dd-3b89768060af.png)
13) Click the “.aspx” file, then click the Page tab on the ribbon above and then the Edit Page button.
![image](https://user-images.githubusercontent.com/84786032/156639155-c1402971-15ae-4553-9c94-3f698dcd3234.png)
15) Click "Add a Web Part" and set Categories to "Media and Content" and Parts to "Content Editor". Click "Add". <br/>
![image](https://user-images.githubusercontent.com/84786032/135382433-71365861-3324-4092-ba33-b8ab077bb93f.png)
12) Click the dropdown arrow in the upper right corner and choose Edit Web Part <br/>
![image](https://user-images.githubusercontent.com/84786032/135382472-74cde243-75bc-4a6b-979b-c265e2aa56d8.png)
14) Under "Content Link" copy-paste link type ./index.html file and under "Appearance" set "Chrome Type" to None. And then click "OK" <br/>
![image](https://user-images.githubusercontent.com/84786032/135657432-f3c7a841-04c3-4df8-8ff8-d09dba15c92a.png)
![image](https://user-images.githubusercontent.com/84786032/156639604-484e748c-bc95-4225-a3f7-3a51d8480470.png)
![image](https://user-images.githubusercontent.com/84786032/156639651-3eaaed7e-dab6-41e9-a52e-d3e18ace514b.png)
![image](https://user-images.githubusercontent.com/84786032/156639741-82765801-8132-457c-b040-9718c479dbb5.png)
16) Click Install on the Installation Required screen <br/>
![image](https://user-images.githubusercontent.com/84786032/135657499-b92ad00a-3b73-41b1-9a31-e729b2ff28d9.png)
18) Click Refresh after installation is loaded <br/>
![image](https://user-images.githubusercontent.com/84786032/135657579-79bc691e-85e6-4c0e-967a-67cffbf16918.png)
20) Click Stop editing if present. 
![image](https://user-images.githubusercontent.com/84786032/135658055-b99a2850-267a-4237-9f48-f3491c4a99aa.png)
22) CONGRATULATIONS!!! ERA is now ready to go! It will appear similar to the screen-shots below. ![image](https://user-images.githubusercontent.com/84786032/156639997-5597e4a0-fcc9-4e3b-9e0a-f4e6b4f11573.png) <br />
<br/> Administrators/Organizers view
![image](https://user-images.githubusercontent.com/84786032/156640035-cf5a5aa9-041b-443a-9685-0a191196e056.png)
<br/> Members/Attendees view
![image](https://user-images.githubusercontent.com/84786032/156643917-7ee77926-8b57-4df9-be30-b18d9b7893d8.png)

# User's Guide
## Organizer
1) Create an event with New Event button and view what security groups are managers and members with Manage Groups
1) View an event's details by clicking the View icon to the left of Title
2) Upload a document by clicking the Upload icon under Documents. View and delete options will be present
3) Edit, delete, and view an event's roster by clicking the Manage Events icon. View Roster will have a print option
3) Type title or date into Search to find a particular event
4) To view coures that have already started or ended click the filters button next to Search and click Show inactive events
### Change Members and Managers group
1) Navaigate to the "Event-Registration" folder and edit the "eventreg-config.json"
2) Enter the name of the new group for either the admin group or the members group
## Attendee
1) Type title or date into Search to find a particular event
2) View an event's details by clicking the View icon to the left of Title
3) Register or unregister by clicking the Registration button
4) If regitered for an event, use the Add event calendar button to add it to Outlook 
