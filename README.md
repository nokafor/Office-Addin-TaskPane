# Office-Addin-TaskPane

This repository is based on the source code used by the [Yo Office generator](https://github.com/OfficeDev/generator-office) when you create a new Office Add-in that appears in the task pane. It is meant as a starting point for developers who are using Script Lab and would like to create a full-fledged Office Add-in.

## Prerequisites

Before using this template, ensure that you have the following installed on your machine:
- [Node.js](https://nodejs.org/en/)
- [Visual Studio Code](https://code.visualstudio.com/download) (optional)
- An Office subscription (you can get a free subscription through the [M365 Developer Program](https://aka.ms/o365devprogram) if you do not have one already)
- a command prompt or terminal

> Command prompts and terminals are normally preinstalled on your machine if you are on Windows or Mac respectively


## Getting started

1. Open a command prompt or terminal and run the following command to download this sample: 

```
git clone https://github.com/nokafor/Office-Addin-TaskPane
```

2. Run the following commands to navigate to the project root and install all the required dependencies.

```
cd Office-Addin-Taskpane
npm install
```

3. Once the install is finished, run the following to open the project in VS Code:

```
code .
```

> If you do not have Visual Studio Code installed, open the following project files in your favorite text editor.

4. Copy your Script Lab snippet to clipboard

  ![script lab screenshot](assets/scriptlabcopy.png)
  
  
5. Paste contents into the "script-lab.txt" file in your project.

6. In script-lab.txt, find the "content" section under "style". 

  ![screenshot of content section from Script Lab](assets/scriptlabcontent.png)

7. Copy the styling content and paste at the bottom of taskpane.css in the src/taskpane folder of the project.

8. In script-lab.txt, find the "content" section under "template". Copy the template content. 

9. Navigate to the taskpane.html file in the src/taskpane folder of project. Paste the content between the <main></main> tags, replacing the content already there.

10. In script-lab.txt, find the "content" section under "script". Copy the script content. 

11. Navigate to the taskpane.ts file in the src/taskpane folder of project. Replace the content of the file with the copied script.
At the top of taskpane.ts file, add the following line:

```
var $ = require( "jquery" );
```

12. Run "npm run build" to build the project and ensure there are no errors

13. Run "npm start" to sideload the add-in and ensure it works.



## TypeScript

This template is written using [TypeScript](http://www.typescriptlang.org/). For the JavaScript version of this template, go to [Office-Addin-TaskPane-JS](https://github.com/OfficeDev/Office-Addin-TaskPane-JS).



## Debugging

This template supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API).  If your question is about the Office JavaScript APIs, make sure it's tagged with  [office-js].

## Additional resources

* [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright

Copyright (c) 2019 Microsoft Corporation. All rights reserved.
