# Color me

This project is showing how you can use Angular 4 with Office.js to build an Excel add-in.

## How to run

In order to run the add-in, you'd need to side-load the add-in within Excel application. Below sections describe the side-loading of manifest file in each of the platforms. 

1A. On Windows, follow [this tutorial](https://dev.office.com/docs/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

1B. On macOS, move the manifest file `angular-office-add-in-manifest.xml` to the folder `/Users/{username}/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (if not exist, create one)

1C. Use the upload my add-in button from the add-in command dialog to upload the manifest file. 

2. Run `ng serve` for a dev server.

3. Open Excel and click this Add-in to load.

<img width="1438" alt="screenshot" src="https://cloud.githubusercontent.com/assets/3375461/25642142/c441e1ea-2f4c-11e7-81a8-d0390b419624.png">

## How to create a new project by yourself

Two parts:

- Generate the Angular project using [Angular CLI](https://github.com/angular/angular-cli).

- Generate the manifest file using [YO Office](https://github.com/OfficeDev/generator-office). When you generate, choose only generating the manifest file.
Then replace all the ports in the generated manifest file from `3000` to `4200`.
